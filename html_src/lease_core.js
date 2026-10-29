// 縣有地租約連動：比對與同步核心（與 ../lease_sync.py 同一套規則）
// 瀏覽器內以全域 LeaseCore 使用；Node 測試時以 require 載入。
(function (root) {
  'use strict';

  const TENANT_SHEET = '承租人總表';   // 工作表名稱開頭即可（實際為「承租人總表1572」）
  const REGISTER_SHEETS = 5;           // 清冊前 5 個工作表是各所清冊
  const CHAR_FIX = { '鲤': '鯉', '台': '臺' };   // 異體字／錯字對照
  // 原檔已大量使用黃、綠、紅底，標記改用原檔沒用過的淺紫
  const MARK = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9B3FF' } };
  const HEAD = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };

  // ---------- 正規化 ----------

  const fixChars = s => s.replace(/[鲤台]/g, c => CHAR_FIX[c]);

  // ExcelJS 的儲存格值可能是公式、富文字、超連結物件，統一取出顯示值
  function cellValue(v) {
    if (v === null || v === undefined) return null;
    if (typeof v !== 'object' || v instanceof Date) return v;
    if ('result' in v) return cellValue(v.result);
    if (v.richText) return v.richText.map(r => r.text).join('');
    if ('text' in v) return v.text;
    if ('error' in v) return null;
    return null;
  }

  const str = v => (v === null || v === undefined) ? '' : String(v);

  const normTown = v => fixChars(str(v).replace(/\s/g, '')).replace(/[市鄉鎮]$/, '');
  const normSection = v => fixChars(str(v).replace(/\s/g, '')).replace(/段$/, '');

  // 1732、'1732-0000'、'6-7內' → '1732-0000' / '0006-0007'；無法解析回傳 null
  function normLot(v) {
    if (v === null || v === undefined) return null;
    const m = str(v).replace(/[\s內]/g, '').match(/^(\d+)(?:-(\d+))?$/);
    if (!m) return null;
    const pad = n => String(parseInt(n || '0', 10)).padStart(4, '0');
    return pad(m[1]) + '-' + pad(m[2]);
  }

  function landKey(town, section, subsection, lot) {
    return [normTown(town), normSection(section), normSection(subsection), normLot(lot)];
  }
  const keyId = k => k.join('|');
  const keyText = ([town, sec, sub, lot]) => `${town} ${sec}段${sub ? ' ' + sub + '小段' : ''} ${lot || '（地號無法辨識）'}`;

  // 去掉括號裡的寄送／代理說明：'王大明(代表繼承:王小明代)' → '王大明'
  const cleanName = n => str(n).replace(/[（(][^)）]*[)）]/g, '').trim();

  const NAME_NOISE = /[（(][^)）]*[)）]|承租|租用|公同共有|等?\d+人|[^一-鿿]/g;
  // 姓名比對用：只留中文字並排序，忽略「、」「等2人」「承租」等寫法差異
  const nameSignature = s => Array.from(fixChars(str(s).replace(NAME_NOISE, ''))).sort().join('');

  const PERIOD_RE = /(\d{2,3})\.(\d{1,2})\.(\d{1,2})\s*[-.~～]\s*(\d{2,3})\.(\d{1,2})\.(\d{1,2})/g;

  // 抓出字串裡所有租期，統一成 '起年,月,日,迄年,月,日' 字串的 Set
  function periodsIn(s) {
    const out = new Set();
    for (const m of str(s).matchAll(PERIOD_RE)) out.add(m.slice(1, 7).map(Number).join(','));
    return out;
  }
  const p2 = n => String(n).padStart(2, '0');
  function fmtPeriod(p) {
    const [a, b, c, d, e, f] = p.split(',').map(Number);
    return `${a}.${p2(b)}.${p2(c)}-${d}.${p2(e)}.${p2(f)}`;
  }
  const endOf = p => { const x = p.split(',').map(Number); return x[3] * 10000 + x[4] * 100 + x[5]; };
  // 租期集合中最晚的到期日（數字 yyyMMdd）；沒有租期回傳 null
  const latestEnd = ps => ps.size ? Math.max(...[...ps].map(endOf)) : null;

  // 欄位內容是否只有租期（可帶承租人姓名，但沒有其他說明文字），只有這種才允許自動覆寫
  function onlyPeriods(s, names) {
    let rest = str(s).replace(PERIOD_RE, '');
    for (const n of names) if (n) rest = rest.split(n).join('');
    return rest.replace(/[\s、，,；;]|等?\d+人/g, '') === '';
  }

  // ---------- 讀檔 ----------

  const orderedSheets = wb => wb.worksheets.slice().sort((a, b) => a.orderNo - b.orderNo);

  // 在前幾列找表頭，回傳 {欄名: [列, 欄]}；合併儲存格只看左上角那格
  function findHeader(ws, wanted, maxRow = 5) {
    const found = {};
    for (let r = 1; r <= maxRow; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 40; c++) {
        const cell = row.getCell(c);
        if (cell.isMerged && cell.master !== cell) continue;
        const v = str(cellValue(cell.value)).replace(/\s/g, '');
        if (wanted.includes(v) && !(v in found)) found[v] = [r, c];
      }
    }
    return found;
  }

  function readTenants(wb) {
    const ws = orderedSheets(wb).find(s => s.name.startsWith(TENANT_SHEET));
    if (!ws) throw new Error(`承租人一覽表找不到「${TENANT_SHEET}…」工作表，是不是選錯檔案？`);
    const col = {};
    ws.getRow(1).eachCell({ includeEmpty: true }, (cell, c) => {
      const h = str(cellValue(cell.value)).replace(/\s/g, '');
      if (!(h in col)) col[h] = c;   // 「地號」出現兩次，取第一個（含 -0000 的那欄）
    });
    const need = ['承租人', '鄉鎮', '段名', '小段', '地號', '租約期間', '承租範圍', '租約到期未續'];
    const missing = need.filter(n => !(n in col));
    if (missing.length) throw new Error('承租人總表缺少欄位：' + missing.join('、'));
    const tenants = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const g = n => cellValue(row.getCell(col[n]).value);
      if (!g('承租人')) continue;
      tenants.push({
        row: r,
        key: landKey(g('鄉鎮'), g('段名'), g('小段'), g('地號')),
        rawLot: g('地號'),
        name: str(g('承租人')).trim(),
        period: str(g('租約期間')).trim(),
        periodCol: col['租約期間'],
        expired: str(g('租約到期未續')),
      });
    }
    return tenants;
  }

  // 各所清冊欄位位置不一樣（苗栗所多一欄、竹南所備註和租約起訖對調），用表頭名稱找欄
  function registerLayout(ws) {
    const h = findHeader(ws, ['鄉鎮市', '鄉鎮', '段', '小段', '地號', '使用關係', '使用關係人', '租約起訖', '備註']);
    if (!h['地號']) return null;
    const lay = {};
    for (const k in h) lay[k] = h[k][1];
    lay.town = lay['鄉鎮市'] || lay['鄉鎮'];
    lay.dataRow = h['地號'][0] + 1;
    return lay;
  }

  function readRegister(wb) {
    const parcels = [];
    for (const ws of orderedSheets(wb).slice(0, REGISTER_SHEETS)) {
      const lay = registerLayout(ws);
      if (!lay) continue;
      for (let r = lay.dataRow; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const g = c => c ? cellValue(row.getCell(c).value) : null;
        const lot = g(lay['地號']);
        if (str(lot).trim() === '' || str(lot).trim() === '地號') continue;
        parcels.push({
          ws, row: r,
          key: landKey(g(lay.town), g(lay['段']), g(lay['小段']), lot),
          rawLot: lot,
          rel: str(g(lay['使用關係'])).trim(),
          whoCell: row.getCell(lay['使用關係人']),
          periodCell: lay['租約起訖'] ? row.getCell(lay['租約起訖']) : null,
        });
      }
    }
    return parcels;
  }

  // ---------- 比對 ----------

  // 由同一地號的所有租約組出清冊應有的「使用關係人」「租約起訖」
  function expectedValues(ts) {
    const names = [...new Set(ts.map(t => cleanName(t.name)))];
    const perName = new Map();
    for (const t of ts) {
      const ps = periodsIn(t.period);
      const txt = ps.size ? fmtPeriod([...ps][0]) : t.period;
      const n = cleanName(t.name);
      if (!perName.has(n)) perName.set(n, txt);
    }
    const distinct = new Set(perName.values());
    const period = distinct.size === 1 ? [...distinct][0]
      : [...perName].map(([n, p]) => n + p).join('、');
    return [names.join('、'), period];
  }

  const isSubset = (a, b) => [...a].every(x => b.has(x));

  // 兩檔是不同年度分開維護的（承租表 114 年版、清冊 115 年版），各自都有對方沒有的續約或換人，
  // 所以逐筆依「哪邊租期較新」決定更新方向，不能無條件以某一邊為準。
  function compare(tenants, parcels) {
    const byKey = new Map();
    for (const t of tenants) {
      const id = keyId(t.key);
      if (!byKey.has(id)) byKey.set(id, []);
      byKey.get(id).push(t);
    }
    const regKeys = new Set(parcels.map(p => keyId(p.key)));
    const res = { same: [], auto: [], reverse: [], manual: [], rentNoTenant: [], tenantNoLand: [], tenantNotRent: [] };

    for (const p of parcels) {
      const ts = byKey.get(keyId(p.key));
      if (!ts) {
        if (p.rel.includes('租')) res.rentNoTenant.push(p);
        continue;
      }
      if (!p.rel.includes('租')) { res.tenantNotRent.push([p, ts]); continue; }
      const whoNow = str(cellValue(p.whoCell.value)).trim();
      const perNow = p.periodCell ? str(cellValue(p.periodCell.value)).trim() : '';
      const [whoNew, perNew] = expectedValues(ts);
      const whoOk = nameSignature(whoNow) === nameSignature(whoNew);
      const want = new Set(ts.flatMap(t => [...periodsIn(t.period)]));
      const have = periodsIn(perNow);
      const perOk = !p.periodCell || !want.size || isSubset(want, have);
      if (whoOk && perOk) { res.same.push(p); continue; }
      const item = { p, ts, whoNow, perNow, whoNew: whoOk ? null : whoNew, perNew: perOk ? null : perNew };

      const reasons = [];
      if (p.rel !== '租用') reasons.push(`使用關係為「${p.rel}」，欄位可能含占用資料`);
      if (item.whoNew && /占|部分|：|:/.test(whoNow)) reasons.push('使用關係人含說明文字');
      const names = ts.map(t => cleanName(t.name)).concat(whoNow.split(/[、，,]/));
      if (item.perNew && perNow && !onlyPeriods(perNow, names)) reasons.push('租約起訖欄含非租期文字');

      const tEnd = latestEnd(want), rEnd = latestEnd(have);
      let newer;
      if (perOk) newer = 'same';
      else if (rEnd === null) newer = 'tenant';
      else if (tEnd === null) newer = 'register';
      else newer = tEnd > rEnd ? 'tenant' : rEnd > tEnd ? 'register' : 'same';
      item.newer = newer;

      const expired = ts.map(t => t.expired).filter(Boolean);
      if (expired.length && newer !== 'register') reasons.push(`承租表註記「${expired[0]}」`);

      if (!reasons.length) {
        if (newer === 'register') {
          // 清冊的租期比較新（已續約）→ 不動清冊，建議回寫承租表
          item.reason = '清冊租期較新，應回寫承租人一覽表';
          item.suggestTenant = [...have].reduce((a, b) => endOf(b) >= endOf(a) ? b : a);
          res.reverse.push(item);
          continue;
        } else if (item.whoNew && whoNow && newer !== 'tenant') {
          reasons.push('姓名不同，租期看不出哪邊較新（繼承／換約／錯字／地號填錯）');
        } else if (newer === 'same' && item.perNew) {
          reasons.push('兩邊到期日相同但起日不同');
        }
      }
      item.reason = reasons.join('；');
      (reasons.length ? res.manual : res.auto).push(item);
    }

    // 承租表有、清冊找不到的地號：用承租人姓名去「清冊租用但無租約」裡找可能的對應（多半是地號打錯）
    const orphanByName = new Map();
    for (const p of res.rentNoTenant) {
      const s = nameSignature(cellValue(p.whoCell.value));
      if (!orphanByName.has(s)) orphanByName.set(s, []);
      orphanByName.get(s).push(p);
    }
    for (const [id, ts] of byKey) {
      if (regKeys.has(id)) continue;
      const hints = new Set();
      for (const t of ts)
        for (const q of orphanByName.get(nameSignature(cleanName(t.name))) || [])
          hints.add(`${q.ws.name} 第${q.row}列 ${q.key[3] ? keyText(q.key) : q.rawLot}`);
      res.tenantNoLand.push([ts[0].key, ts, [...hints].join('；')]);
    }
    return res;
  }

  // ---------- 同步 ----------

  function mark(cell, value, old, today, source) {
    cell.value = value;
    // 載入的檔案裡多格會共用同一個樣式物件，複製一份再改，避免整片一起變色
    cell.style = Object.assign({}, cell.style, { fill: MARK });
    cell.note = `${today} 依${source}同步\n原值：${old || '（空白）'}`;
  }

  function applySync(res, today) {
    let n = 0;
    for (const it of res.auto) {
      for (const [cell, nv, old] of [[it.p.whoCell, it.whoNew, it.whoNow], [it.p.periodCell, it.perNew, it.perNow]]) {
        if (nv === null || !cell) continue;
        mark(cell, nv, old, today, '承租人一覽表');
        n++;
      }
    }
    return n;
  }

  // 只有單一承租人、且兩邊姓名相同時才自動回寫租期：
  // 多人共用一筆地號時各人租期可能不同；姓名不同代表續約時換了人，承租表的身分證、地址也要跟著換
  const reverseOk = it => new Set(it.ts.map(t => cleanName(t.name))).size === 1 && it.whoNew === null;
  function reverseNote(it) {
    if (reverseOk(it)) return '是';
    if (it.whoNew !== null) return '否（承租人已換人，需人工改姓名／身分證／地址）';
    return '否（多位承租人，請人工改）';
  }

  function applyReverse(res, tenantWb, today) {
    const ws = orderedSheets(tenantWb).find(s => s.name.startsWith(TENANT_SHEET));
    const done = new Set();
    for (const it of res.reverse) {
      if (!reverseOk(it)) continue;
      const nv = fmtPeriod(it.suggestTenant);
      for (const t of it.ts) {
        if (done.has(t.row)) continue;
        mark(ws.getRow(t.row).getCell(t.periodCol), nv, t.period, today, '縣有地清冊');
        done.add(t.row);
      }
    }
    return done.size;
  }

  // ---------- 差異報告 ----------

  function addSheet(wb, title, header, rows, widths) {
    const ws = wb.addWorksheet(title, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.addRow(header);
    ws.getRow(1).eachCell(c => { c.font = { bold: true }; c.fill = HEAD; });
    for (const r of rows) ws.addRow(r.map(v => v === null || v === undefined ? '' : v));
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    for (let r = 2; r <= ws.rowCount; r++)
      ws.getRow(r).alignment = { wrapText: true, vertical: 'top' };
    if (rows.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length + 1, column: header.length } };
    return ws;
  }

  const tenantDesc = ts => ts.map(t => `第${t.row}列 ${t.name} ${t.period}${t.expired ? ' ' + t.expired : ''}`).join('\n');

  function summaryRows(res) {
    return [
      ['已一致', res.same.length],
      ['① 承租表較新 → 更新清冊（同步時會改，紫底標示）', res.auto.length],
      ['② 清冊較新 → 回寫承租表（同步時單一承租人者會改）', res.reverse.length],
      ['③ 需人工確認（不會自動改）', res.manual.length],
      ['④ 清冊標租用，承租人表查無租約', res.rentNoTenant.length],
      ['⑤ 承租人表有租約，清冊找不到該地號', res.tenantNoLand.length],
      ['⑥ 承租人表有租約，清冊使用關係不是租用', res.tenantNotRent.length],
    ];
  }

  function buildReport(ExcelJS, res, info) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('總覽');
    ws.addRows([
      ['縣有地租約連動 差異報告', ''],
      ['產生時間', info.time],
      ['承租人一覽表', info.tenantFile],
      ['清冊', info.registerFile],
      ['同步後檔案', info.synced.length ? info.synced.join('\n') : '（未執行同步，只比對）'],
      ['', ''],
      ['項目', '筆數'],
      ...summaryRows(res),
    ]);
    ws.getCell('A1').font = { bold: true, size: 14 };
    ws.getRow(7).eachCell(c => { c.font = { bold: true }; c.fill = HEAD; });
    ws.getColumn(1).width = 46;
    ws.getColumn(2).width = 60;
    ws.getCell('B5').alignment = { wrapText: true };

    const diffRows = items => items.map(it => [
      it.p.ws.name, it.p.row, keyText(it.p.key), it.p.rel,
      it.whoNow, it.whoNew || '（不變）', it.perNow, it.perNew || '（不變）',
      tenantDesc(it.ts), it.reason || '']);
    const head = ['清冊工作表', '清冊列', '地號', '使用關係', '使用關係人(現在)', '使用關係人(承租表)',
      '租約起訖(現在)', '租約起訖(承租表)', '承租人表對應列', '不自動改的原因'];
    const wid = [22, 7, 26, 11, 30, 30, 28, 28, 40, 30];

    addSheet(wb, '①承租表較新→改清冊', head.slice(0, -1), diffRows(res.auto).map(r => r.slice(0, -1)), wid.slice(0, -1));
    addSheet(wb, '②清冊較新→改承租表',
      ['清冊工作表', '清冊列', '地號', '使用關係人(清冊)', '租約起訖(清冊)', '承租人表對應列(舊)', '建議承租表租約期間', '同步時是否自動改'],
      res.reverse.map(it => [it.p.ws.name, it.p.row, keyText(it.p.key), it.whoNow, it.perNow,
        tenantDesc(it.ts), fmtPeriod(it.suggestTenant), reverseNote(it)]),
      [22, 7, 26, 24, 28, 40, 22, 24]);
    addSheet(wb, '③需人工確認', head, diffRows(res.manual), wid);
    addSheet(wb, '④清冊租用但無租約',
      ['清冊工作表', '清冊列', '地號(原寫法)', '地號', '使用關係', '使用關係人', '租約起訖'],
      res.rentNoTenant.map(p => [p.ws.name, p.row, str(p.rawLot), keyText(p.key), p.rel,
        str(cellValue(p.whoCell.value)), p.periodCell ? str(cellValue(p.periodCell.value)) : '']),
      [22, 7, 14, 26, 11, 30, 28]);
    addSheet(wb, '⑤承租表地號清冊查無',
      ['地號(正規化)', '承租人表列', '承租人', '地號(原寫法)', '租約期間', '可能原因', '清冊中同名承租人（可能是同一筆）'],
      res.tenantNoLand.flatMap(([k, ts, hint]) => ts.map(t => [keyText(k), t.row, t.name, str(t.rawLot), t.period,
        k[0].includes('段') ? '鄉鎮欄填錯' : '清冊無此地號，或段名／地號打錯', hint])),
      [26, 9, 20, 14, 24, 26, 40]);
    addSheet(wb, '⑥承租表有約但清冊非租用',
      ['清冊工作表', '清冊列', '地號', '清冊使用關係', '清冊使用關係人', '承租人表對應列'],
      res.tenantNotRent.map(([p, ts]) => [p.ws.name, p.row, keyText(p.key), p.rel || '（空白）',
        str(cellValue(p.whoCell.value)), tenantDesc(ts)]),
      [22, 7, 26, 12, 30, 40]);
    return wb;
  }

  // ---------- 主流程 ----------

  const pad2 = n => String(n).padStart(2, '0');
  function stamps(d = new Date()) {
    const day = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    return {
      today: day,
      time: `${day} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`,
      file: `${day.replace(/-/g, '')}_${pad2(d.getHours())}${pad2(d.getMinutes())}`,
    };
  }
  const baseName = f => f.replace(/\.xlsx$/i, '');

  // tenantBuf / registerBuf：ArrayBuffer 或 Buffer；回傳 {res, outputs:[{name, workbook}]}
  async function run(ExcelJS, tenantBuf, tenantName, registerBuf, registerName, doSync, log = () => {}) {
    const st = stamps();
    log('讀取承租人一覽表…');
    const twb = new ExcelJS.Workbook();
    await twb.xlsx.load(tenantBuf);
    const tenants = readTenants(twb);
    log(`  租約 ${tenants.length} 筆`);
    log('讀取縣有地清冊…');
    const rwb = new ExcelJS.Workbook();
    await rwb.xlsx.load(registerBuf);
    const parcels = readRegister(rwb);
    if (!parcels.length) throw new Error('清冊前 5 個工作表都找不到「地號」表頭，是不是選錯檔案？');
    log(`  地號 ${parcels.length} 筆`);
    const res = compare(tenants, parcels);

    const outputs = [];
    if (doSync) {
      const n1 = applySync(res, st.today);
      outputs.push({ name: `${baseName(registerName)}_已同步_${st.file}.xlsx`, workbook: rwb, note: `清冊更新 ${n1} 格` });
      const n2 = applyReverse(res, twb, st.today);
      outputs.push({ name: `${baseName(tenantName)}_已同步_${st.file}.xlsx`, workbook: twb, note: `承租表更新 ${n2} 格` });
    }
    const report = buildReport(ExcelJS, res, {
      time: st.time, tenantFile: tenantName, registerFile: registerName, synced: outputs.map(o => o.name),
    });
    outputs.unshift({ name: `差異報告_${st.file}.xlsx`, workbook: report, note: '差異報告' });
    return { res, summary: summaryRows(res), outputs };
  }

  const api = { run, compare, normLot, normTown, normSection, nameSignature, periodsIn, onlyPeriods };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LeaseCore = api;
})(typeof self !== 'undefined' ? self : this);
