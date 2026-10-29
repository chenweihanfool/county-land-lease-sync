# -*- coding: utf-8 -*-
"""縣有地租約連動工具

依地號（鄉鎮＋段＋小段＋地號）比對「承租人一覽表」與「縣有農牧用地清冊」，
以租期較新的一邊為準，把承租人姓名與租約期間同步到另一邊：
  承租表較新 → 更新清冊「使用關係人」「租約起訖」
  清冊較新   → 回寫承租表「租約期間」
無法判斷的（姓名不同但租期相同、欄位夾帶占用說明等）只列報告，不自動改。

用法：
    python lease_sync.py 承租人一覽表.xlsx 縣有地清冊.xlsx            # 只比對，產生差異報告
    python lease_sync.py 承租人一覽表.xlsx 縣有地清冊.xlsx --sync     # 比對＋另存同步後清冊
    python lease_sync.py                                              # 不帶參數：跳出視窗選檔

原始檔案永遠不會被修改；輸出檔放在清冊同一個資料夾。
"""
import argparse
import collections
import datetime
import os
import re
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TENANT_SHEET = '承租人總表'          # 工作表名稱開頭即可（實際為「承租人總表1572」）
REGISTER_SHEETS = 5                  # 清冊前 5 個工作表是各所清冊

# 異體字／錯字對照：鍵值比對前先換掉（例如承租表「鲤魚段」用了簡體字）
CHAR_FIX = str.maketrans({'鲤': '鯉', '台': '臺'})

MARK = PatternFill('solid', fgColor='FFD9B3FF')   # 原檔已大量使用黃、綠、紅底，標記改用原檔沒用過的淺紫
HEAD = PatternFill('solid', fgColor='DDEBF7')


# ---------- 正規化 ----------

def norm_town(v):
    s = re.sub(r'\s', '', str(v or '')).translate(CHAR_FIX)
    return re.sub(r'[市鄉鎮]$', '', s)


def norm_section(v):
    s = re.sub(r'\s', '', str(v or '')).translate(CHAR_FIX)
    return re.sub(r'段$', '', s)


def norm_lot(v):
    """1732、'1732-0000'、'6-7內'、1732.0 → '1732-0000' / '0006-0007'；無法解析回傳 None。"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r'[\s內]', '', str(v))
    m = re.fullmatch(r'(\d+)(?:-(\d+))?', s)
    if not m:
        return None
    return f'{int(m.group(1)):04d}-{int(m.group(2) or 0):04d}'


def land_key(town, section, subsection, lot):
    return (norm_town(town), norm_section(section), norm_section(subsection), norm_lot(lot))


def key_text(k):
    town, sec, sub, lot = k
    return f"{town} {sec}段{(' ' + sub + '小段') if sub else ''} {lot or '（地號無法辨識）'}"


def clean_name(name):
    """去掉括號裡的寄送／代理說明：'王大明(代表繼承:王小明代)' → '王大明'。"""
    return re.sub(r'[（(][^)）]*[)）]', '', str(name or '')).strip()


NAME_NOISE = re.compile(r'[（(][^)）]*[)）]|承租|租用|公同共有|等?\d+人|[^一-鿿]')


def name_signature(s):
    """姓名比對用：只留中文字並排序，忽略「、」「等2人」「承租」等寫法差異。"""
    return ''.join(sorted(NAME_NOISE.sub('', str(s or '')).translate(CHAR_FIX)))


PERIOD_RE = re.compile(r'(\d{2,3})\.(\d{1,2})\.(\d{1,2})\s*[-.~～]\s*(\d{2,3})\.(\d{1,2})\.(\d{1,2})')


def periods_in(s):
    """抓出字串裡所有租期，統一成 (起年,月,日,迄年,月,日)。"""
    return {tuple(int(x) for x in m.groups()) for m in PERIOD_RE.finditer(str(s or ''))}


def fmt_period(p):
    return '{}.{:02d}.{:02d}-{}.{:02d}.{:02d}'.format(*p)


def only_periods(s, names):
    """欄位內容是否只有租期（可帶承租人姓名，但沒有其他說明文字），只有這種才允許自動覆寫。"""
    rest = PERIOD_RE.sub('', str(s or ''))
    for n in names:
        rest = rest.replace(n, '')
    return re.sub(r'[\s、，,；;]|等?\d+人', '', rest) == ''


# ---------- 讀檔 ----------

def find_header(ws, wanted, max_row=5):
    """在前幾列找表頭，回傳 {欄名: 欄索引}，同時看上下兩層表頭。"""
    found = {}
    for r in range(1, max_row + 1):
        for c in ws[r][:40]:
            v = re.sub(r'\s', '', str(c.value or ''))
            if v in wanted and v not in found:
                found[v] = (r, c.column - 1)
    return found


def read_tenants(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = next(s for s in wb.worksheets if s.title.startswith(TENANT_SHEET))
    rows = list(ws.iter_rows(values_only=True))
    head = [re.sub(r'\s', '', str(v or '')) for v in rows[0]]
    col = {}
    for i, h in enumerate(head):
        col.setdefault(h, i)   # 「地號」出現兩次，取第一個（含 -0000 的那欄）
    need = ['承租人', '鄉鎮', '段名', '小段', '地號', '租約期間', '承租範圍', '租約到期未續']
    missing = [n for n in need if n not in col]
    if missing:
        raise SystemExit(f'承租人總表缺少欄位：{missing}')
    tenants = []
    for n, r in enumerate(rows[1:], start=2):
        if not r[col['承租人']]:
            continue
        tenants.append(dict(
            row=n,
            key=land_key(r[col['鄉鎮']], r[col['段名']], r[col['小段']], r[col['地號']]),
            raw_lot=r[col['地號']],
            name=str(r[col['承租人']]).strip(),
            period=str(r[col['租約期間']] or '').strip(),
            period_col=col['租約期間'] + 1,
            scope=r[col['承租範圍']] or '',
            expired=r[col['租約到期未續']] or '',
        ))
    wb.close()
    return tenants


def register_layout(ws):
    """各所清冊欄位位置不一樣（苗栗所多一欄、竹南所備註和租約起訖對調），用表頭名稱找欄。"""
    h = find_header(ws, {'鄉鎮市', '鄉鎮', '段', '小段', '地號', '使用關係', '使用關係人', '租約起訖', '備註'})
    if '地號' not in h:
        return None
    lay = {k: v[1] for k, v in h.items()}
    lay['town'] = lay.get('鄉鎮市', lay.get('鄉鎮'))
    lay['data_row'] = h['地號'][0] + 1
    return lay


def read_register(wb):
    parcels = []
    for ws in wb.worksheets[:REGISTER_SHEETS]:
        lay = register_layout(ws)
        if not lay:
            continue
        for row in ws.iter_rows(min_row=lay['data_row']):
            lot = row[lay['地號']].value
            if lot in (None, '', ' ') or str(lot).strip() in ('地號',):
                continue
            k = land_key(row[lay['town']].value, row[lay['段']].value,
                         row[lay['小段']].value if '小段' in lay else None, lot)
            parcels.append(dict(
                ws=ws, row=row[0].row, key=k, raw_lot=lot,
                rel=str(row[lay['使用關係']].value or '').strip(),
                who_cell=row[lay['使用關係人']],
                period_cell=row[lay['租約起訖']] if '租約起訖' in lay else None,
            ))
    return parcels


# ---------- 比對 ----------

def expected_values(ts):
    """由同一地號的所有租約組出清冊應有的「使用關係人」「租約起訖」。"""
    names = list(dict.fromkeys(clean_name(t['name']) for t in ts))
    who = '、'.join(names)
    per_name = collections.OrderedDict()
    for t in ts:
        ps = periods_in(t['period'])
        txt = fmt_period(next(iter(ps))) if ps else t['period']
        per_name.setdefault(clean_name(t['name']), txt)
    distinct = set(per_name.values())
    if len(distinct) == 1:
        period = distinct.pop()
    else:
        period = '、'.join(f'{n}{p}' for n, p in per_name.items())
    return who, period


def latest_end(periods):
    """租期集合中最晚的到期日 (年,月,日)；沒有租期回傳 None。"""
    return max((p[3:] for p in periods), default=None)


def compare(tenants, parcels):
    """逐筆清冊地號比對，依「哪邊租期較新」決定更新方向。

    兩檔是不同年度分開維護的（承租表 114 年版、清冊 115 年版），
    各自都有對方沒有的續約或換人，所以不能無條件以某一邊為準。
    """
    by_key = collections.defaultdict(list)
    for t in tenants:
        by_key[t['key']].append(t)
    reg_keys = {p['key'] for p in parcels}

    result = dict(same=[], auto=[], reverse=[], manual=[],
                  rent_no_tenant=[], tenant_no_land=[], tenant_not_rent=[])
    for p in parcels:
        ts = by_key.get(p['key'])
        if not ts:
            if '租' in p['rel']:
                result['rent_no_tenant'].append(p)
            continue
        if '租' not in p['rel']:
            result['tenant_not_rent'].append((p, ts))
            continue
        who_now = str(p['who_cell'].value or '').strip()
        per_now = str(p['period_cell'].value or '').strip() if p['period_cell'] else ''
        who_new, per_new = expected_values(ts)
        who_ok = name_signature(who_now) == name_signature(who_new)
        want = {q for t in ts for q in periods_in(t['period'])}
        have = periods_in(per_now)
        per_ok = p['period_cell'] is None or not want or want <= have
        if who_ok and per_ok:
            result['same'].append(p)
            continue
        item = dict(p=p, ts=ts, who_now=who_now, per_now=per_now,
                    who_new=None if who_ok else who_new,
                    per_new=None if per_ok else per_new)

        reasons = []
        if p['rel'] != '租用':
            reasons.append(f'使用關係為「{p["rel"]}」，欄位可能含占用資料')
        if item['who_new'] and re.search(r'占|部分|：|:', who_now):
            reasons.append('使用關係人含說明文字')
        names = [clean_name(t['name']) for t in ts] + re.split(r'[、，,]', who_now)
        if item['per_new'] and per_now and not only_periods(per_now, names):
            reasons.append('租約起訖欄含非租期文字')

        t_end, r_end = latest_end(want), latest_end(have)
        if per_ok:
            newer = 'same'
        elif r_end is None:
            newer = 'tenant'
        elif t_end is None:
            newer = 'register'
        else:
            newer = 'tenant' if t_end > r_end else 'register' if r_end > t_end else 'same'
        item['newer'] = newer

        expired = [t['expired'] for t in ts if t['expired']]
        if expired and newer != 'register':
            reasons.append(f'承租表註記「{expired[0]}」')

        if reasons:
            pass
        elif newer == 'register':
            # 清冊的租期比較新（已續約）→ 不動清冊，建議回寫承租表
            reasons.append('清冊租期較新，應回寫承租人一覽表')
            item['suggest_tenant'] = sorted(have, key=lambda q: q[3:])[-1]
            result['reverse'].append(item)
            item['reason'] = reasons[0]
            continue
        elif item['who_new'] and who_now and newer != 'tenant':
            reasons.append('姓名不同，租期看不出哪邊較新（繼承／換約／錯字／地號填錯）')
        elif newer == 'same' and item['per_new']:
            reasons.append('兩邊到期日相同但起日不同')
        item['reason'] = '；'.join(reasons)
        result['manual' if reasons else 'auto'].append(item)

    # 承租表有、清冊找不到的地號：用承租人姓名去「清冊租用但無租約」裡找可能的對應（多半是地號打錯）
    orphan_by_name = collections.defaultdict(list)
    for p in result['rent_no_tenant']:
        orphan_by_name[name_signature(p['who_cell'].value)].append(p)
    for k, ts in by_key.items():
        if k not in reg_keys:
            hints = [f"{q['ws'].title} 第{q['row']}列 {key_text(q['key']) if q['key'][3] else q['raw_lot']}"
                     for t in ts for q in orphan_by_name.get(name_signature(clean_name(t['name'])), [])]
            result['tenant_no_land'].append((k, ts, '；'.join(dict.fromkeys(hints))))
    return result


# ---------- 輸出 ----------

def write_sheet(wb, title, header, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = HEAD
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.freeze_panes = 'A2'
    if rows:
        ws.auto_filter.ref = ws.dimensions
    return ws


def tenant_desc(ts):
    return '\n'.join(f"第{t['row']}列 {t['name']} {t['period']}{' ' + t['expired'] if t['expired'] else ''}" for t in ts)


def write_report(res, path, tenant_file, register_file, synced_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '總覽'
    lines = [
        ('縣有地租約連動 差異報告', ''),
        ('產生時間', datetime.datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('主檔（承租人一覽表）', os.path.basename(tenant_file)),
        ('清冊', os.path.basename(register_file)),
        ('同步後檔案', '\n'.join(os.path.basename(f) for f in synced_file) if synced_file else '（未執行同步，只比對）'),
        ('', ''),
        ('項目', '筆數'),
        ('已一致', len(res['same'])),
        ('① 承租表較新 → 更新清冊（同步時會改，紫底標示）', len(res['auto'])),
        ('② 清冊較新 → 回寫承租表（同步時單一承租人者會改）', len(res['reverse'])),
        ('③ 需人工確認（不會自動改）', len(res['manual'])),
        ('④ 清冊標租用，承租人表查無租約', len(res['rent_no_tenant'])),
        ('⑤ 承租人表有租約，清冊找不到該地號', len(res['tenant_no_land'])),
        ('⑥ 承租人表有租約，清冊使用關係不是租用', len(res['tenant_not_rent'])),
    ]
    for l in lines:
        ws.append(l)
    ws['A1'].font = Font(bold=True, size=14)
    for c in ws[7]:
        c.font = Font(bold=True)
        c.fill = HEAD
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 60

    def diff_rows(items):
        out = []
        for it in items:
            p = it['p']
            out.append([p['ws'].title, p['row'], key_text(p['key']), p['rel'],
                        it['who_now'], it['who_new'] or '（不變）',
                        it['per_now'], it['per_new'] or '（不變）',
                        tenant_desc(it['ts']), it.get('reason', '')])
        return out

    head = ['清冊工作表', '清冊列', '地號', '使用關係', '使用關係人(現在)', '使用關係人(承租表)',
            '租約起訖(現在)', '租約起訖(承租表)', '承租人表對應列', '不自動改的原因']
    wid = [22, 7, 26, 11, 30, 30, 28, 28, 40, 30]
    write_sheet(wb, '①承租表較新→改清冊', head[:-1], [r[:-1] for r in diff_rows(res['auto'])], wid[:-1])
    write_sheet(wb, '②清冊較新→改承租表',
                ['清冊工作表', '清冊列', '地號', '使用關係人(清冊)', '租約起訖(清冊)', '承租人表對應列(舊)',
                 '建議承租表租約期間', '同步時是否自動改'],
                [[it['p']['ws'].title, it['p']['row'], key_text(it['p']['key']), it['who_now'], it['per_now'],
                  tenant_desc(it['ts']), fmt_period(it['suggest_tenant']),
                  reverse_note(it)] for it in res['reverse']],
                [22, 7, 26, 24, 28, 40, 22, 24])
    write_sheet(wb, '③需人工確認', head, diff_rows(res['manual']), wid)
    write_sheet(wb, '④清冊租用但無租約',
                ['清冊工作表', '清冊列', '地號(原寫法)', '地號', '使用關係', '使用關係人', '租約起訖'],
                [[p['ws'].title, p['row'], str(p['raw_lot']), key_text(p['key']), p['rel'],
                  p['who_cell'].value, p['period_cell'].value if p['period_cell'] else '']
                 for p in res['rent_no_tenant']], [22, 7, 14, 26, 11, 30, 28])
    write_sheet(wb, '⑤承租表地號清冊查無',
                ['地號(正規化)', '承租人表列', '承租人', '地號(原寫法)', '租約期間', '可能原因', '清冊中同名承租人（可能是同一筆）'],
                [[key_text(k), t['row'], t['name'], str(t['raw_lot']), t['period'],
                  '鄉鎮欄填錯' if '段' in k[0] else '清冊無此地號，或段名／地號打錯', hint]
                 for k, ts, hint in res['tenant_no_land'] for t in ts], [26, 9, 20, 14, 24, 26, 40])
    write_sheet(wb, '⑥承租表有約但清冊非租用',
                ['清冊工作表', '清冊列', '地號', '清冊使用關係', '清冊使用關係人', '承租人表對應列'],
                [[p['ws'].title, p['row'], key_text(p['key']), p['rel'] or '（空白）',
                  p['who_cell'].value, tenant_desc(ts)] for p, ts in res['tenant_not_rent']],
                [22, 7, 26, 12, 30, 40])
    wb.save(path)


def mark(cell, new, old, today, source):
    cell.value = new
    cell.fill = MARK
    cell.comment = Comment(f'{today} 依{source}同步\n原值：{old or "（空白）"}', '連動工具')


def apply_sync(res, today):
    """把①的建議值寫進清冊（記憶體中），改過的格子塗淺紫並把舊值留在註解。"""
    n = 0
    for it in res['auto']:
        for cell, new, old in ((it['p']['who_cell'], it['who_new'], it['who_now']),
                               (it['p']['period_cell'], it['per_new'], it['per_now'])):
            if new is None or cell is None:
                continue
            mark(cell, new, old, today, '承租人一覽表')
            n += 1
    return n


def reverse_ok(it):
    """只有單一承租人、且兩邊姓名相同時才自動回寫租期。

    多人共用一筆地號時各人租期可能不同；姓名不同代表續約時換了承租人，
    承租表那列的身分證、地址也要跟著換，只改租期反而會錯。
    """
    return len({clean_name(t['name']) for t in it['ts']}) == 1 and it['who_new'] is None


def reverse_note(it):
    if reverse_ok(it):
        return '是'
    if it['who_new'] is not None:
        return '否（承租人已換人，需人工改姓名／身分證／地址）'
    return '否（多位承租人，請人工改）'


def apply_reverse(res, tenant_wb, today):
    """把②清冊較新的租期回寫到承租人總表的「租約期間」。"""
    ws = next(s for s in tenant_wb.worksheets if s.title.startswith(TENANT_SHEET))
    done = set()
    for it in res['reverse']:
        if not reverse_ok(it):
            continue
        new = fmt_period(it['suggest_tenant'])
        for t in it['ts']:
            if t['row'] in done:
                continue
            mark(ws.cell(t['row'], t['period_col']), new, t['period'], today, '縣有地清冊')
            done.add(t['row'])
    return len(done)


def run(tenant_file, register_file, do_sync, log=print):
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out_dir = os.path.dirname(os.path.abspath(register_file))
    log('讀取承租人一覽表…')
    tenants = read_tenants(tenant_file)
    log(f'  租約 {len(tenants)} 筆')
    log('讀取縣有地清冊…（檔案大，約需半分鐘）')
    wb = openpyxl.load_workbook(register_file)
    parcels = read_register(wb)
    log(f'  地號 {len(parcels)} 筆')
    res = compare(tenants, parcels)

    synced = []
    if do_sync:
        today = datetime.date.today().isoformat()
        n = apply_sync(res, today)
        base = os.path.splitext(os.path.basename(register_file))[0]
        out = os.path.join(out_dir, f'{base}_已同步_{stamp}.xlsx')
        wb.save(out)
        synced.append(out)
        log(f'清冊更新 {n} 格 → {out}')
        twb = openpyxl.load_workbook(tenant_file)
        n = apply_reverse(res, twb, today)
        base = os.path.splitext(os.path.basename(tenant_file))[0]
        out = os.path.join(out_dir, f'{base}_已同步_{stamp}.xlsx')
        twb.save(out)
        synced.append(out)
        log(f'承租表更新 {n} 格 → {out}')
    report = os.path.join(out_dir, f'差異報告_{stamp}.xlsx')
    write_report(res, report, tenant_file, register_file, synced)
    log(f'差異報告 → {report}')
    log('結果：一致 {} ／①改清冊 {} ／②改承租表 {} ／③人工確認 {} ／④清冊租用無租約 {} ／⑤清冊查無地號 {} ／⑥清冊非租用 {}'.format(
        len(res['same']), len(res['auto']), len(res['reverse']), len(res['manual']),
        len(res['rent_no_tenant']), len(res['tenant_no_land']), len(res['tenant_not_rent'])))
    return report, synced


def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    t = filedialog.askopenfilename(title='① 選「承租人一覽表」', filetypes=[('Excel', '*.xlsx')])
    if not t:
        return
    r = filedialog.askopenfilename(title='② 選「縣有農牧用地清冊」', filetypes=[('Excel', '*.xlsx')])
    if not r:
        return
    sync = messagebox.askyesno('連動工具', '要順便產生「同步後」的檔案嗎？\n\n是：比對＋另存兩份已更新的檔案（原檔不動）\n否：只產生差異報告')
    try:
        report, synced = run(t, r, sync)
    except Exception as e:  # 視窗模式下要讓使用者看得到錯誤
        messagebox.showerror('連動工具', f'發生錯誤：{e}')
        raise
    messagebox.showinfo('連動工具', '完成！\n\n差異報告：' + report + ''.join('\n同步後：' + f for f in synced))
    os.startfile(os.path.dirname(report))


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) == 1:
        return gui()
    ap = argparse.ArgumentParser(description='縣有地租約連動工具')
    ap.add_argument('tenant', help='承租人一覽表 .xlsx')
    ap.add_argument('register', help='縣有農牧用地清冊 .xlsx')
    ap.add_argument('--sync', action='store_true', help='另存同步後的清冊')
    a = ap.parse_args()
    run(a.tenant, a.register, a.sync)


if __name__ == '__main__':
    main()
