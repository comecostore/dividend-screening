"""消えた39社を再取得して復元するスクリプト"""
import json, requests, re, time, sys, io, datetime
from bs4 import BeautifulSoup

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'ja,en-US;q=0.9',
}
TODAY_YR = datetime.date.today().year

def parse_num(s):
    s = str(s).replace(',','').strip()
    s = re.sub(r'[^\d\.\-]','',s)
    try: return float(s) if s not in ('','-') else None
    except: return None

def parse_oku(s):
    s = str(s).strip().replace(',','')
    m = re.search(r'([\-\d\.]+)億', s)
    if m: return round(float(m.group(1)) * 100, 1)
    m2 = re.search(r'([\-\d\.]+)兆', s)
    if m2: return round(float(m2.group(1)) * 100_000, 1)
    return parse_num(s)

def yr_label(cell_text):
    m = re.search(r'(\d{4})[/年](\d{1,2})', cell_text)
    return f"{m.group(1)}/{int(m.group(2)):02d}" if m else None

def find_col(hdr, keywords):
    for i, h in enumerate(hdr):
        if any(k in h for k in keywords):
            return i
    return None

def fetch_results(code):
    try:
        r = requests.get(f'https://irbank.net/{code}/results', headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.content, 'html.parser')
        tables = soup.find_all('table')
        if not tables: return None
        t0 = tables[0]
        rows0 = t0.find_all('tr')
        hdr0 = [c.get_text(strip=True) for c in rows0[0].find_all(['th','td'])]
        sales_i  = find_col(hdr0, ['売上', '営業収益', '経常収益', '純営業収益', '収益',
                                    '正味収入保険料', '営業総収入', '営業収入', '完成工事高'])
        eps_i    = find_col(hdr0, ['EPS'])
        oprate_i = find_col(hdr0, ['営利率'])
        main = {}
        for row in rows0[1:]:
            cells = [c.get_text(strip=True) for c in row.find_all(['th','td'])]
            if not cells or not cells[0]: continue
            yr_m = re.search(r'(\d{4})', cells[0])
            if not yr_m or int(yr_m.group(1)) > TODAY_YR: continue
            lb = yr_label(cells[0]) or yr_m.group(1)
            main[lb] = {
                'sales':     parse_oku(cells[sales_i])  if sales_i  is not None and sales_i  < len(cells) else None,
                'eps':       parse_num(cells[eps_i])    if eps_i    is not None and eps_i    < len(cells) else None,
                'op_margin': parse_num(cells[oprate_i]) if oprate_i is not None and oprate_i < len(cells) else None,
            }
        eq_data = {}
        if len(tables) > 1:
            t1 = tables[1]
            rows1 = t1.find_all('tr')
            hdr1 = [c.get_text(strip=True) for c in rows1[0].find_all(['th','td'])]
            eq_i = find_col(hdr1, ['自己資本比率', '株主資本比率'])
            if eq_i is not None:
                for row in rows1[1:]:
                    cells = [c.get_text(strip=True) for c in row.find_all(['th','td'])]
                    if not cells or not cells[0]: continue
                    yr_m = re.search(r'(\d{4})', cells[0])
                    if not yr_m or int(yr_m.group(1)) > TODAY_YR: continue
                    lb = yr_label(cells[0]) or yr_m.group(1)
                    eq_data[lb] = parse_num(cells[eq_i]) if eq_i < len(cells) else None
        cf_data = {}
        cf_cash_data = {}
        if len(tables) > 2:
            t2 = tables[2]
            rows2 = t2.find_all('tr')
            hdr2 = [c.get_text(strip=True) for c in rows2[0].find_all(['th','td'])]
            ocf_i   = find_col(hdr2, ['営業CF'])
            cash_i2 = find_col(hdr2, ['現金等', '現金及び現金同等物'])
            for row in rows2[1:]:
                cells = [c.get_text(strip=True) for c in row.find_all(['th','td'])]
                if not cells or not cells[0]: continue
                yr_m = re.search(r'(\d{4})', cells[0])
                if not yr_m or int(yr_m.group(1)) > TODAY_YR: continue
                lb = yr_label(cells[0]) or yr_m.group(1)
                if ocf_i is not None:
                    val = parse_oku(cells[ocf_i]) if ocf_i < len(cells) else None
                    if val is not None: cf_data[lb] = val
                if cash_i2 is not None:
                    val2 = parse_oku(cells[cash_i2]) if cash_i2 < len(cells) else None
                    if val2 is not None: cf_cash_data[lb] = val2
        sy = sorted(main.keys())[-10:]
        ey = sorted(eq_data.keys())[-10:]
        cy = sorted(cf_data.keys())[-10:]
        ccy = sorted(cf_cash_data.keys())[-10:]
        return {
            'years': sy, 'sales': [main[y]['sales'] for y in sy],
            'eps': [main[y]['eps'] for y in sy], 'op_margin': [main[y]['op_margin'] for y in sy],
            'eq_years': ey, 'equity': [eq_data[y] for y in ey],
            'cf_years': cy, 'ocf': [cf_data[y] for y in cy],
            'cf_cash_years': ccy, 'cf_cash': [cf_cash_data[y] for y in ccy],
        }
    except Exception as e:
        print(f'  エラー: {e}')
        return None

def fetch_dividend(code):
    try:
        r = requests.get(f'https://irbank.net/{code}/dividend', headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.content, 'html.parser')
        for t in soup.find_all('table'):
            rows = t.find_all('tr')
            if not rows: continue
            hdr = [c.get_text(strip=True) for c in rows[0].find_all(['th','td'])]
            if '合計' not in hdr and '分割調整' not in hdr: continue
            col = hdr.index('分割調整') if '分割調整' in hdr else hdr.index('合計')
            items = []
            for row in rows[1:]:
                cells = [c.get_text(strip=True) for c in row.find_all(['th','td'])]
                if len(cells) <= col: continue
                val = parse_num(cells[col])
                if val is None or val <= 0: continue
                yr_m = re.search(r'(\d{4})年', cells[0])
                if not yr_m: continue
                yr = int(yr_m.group(1))
                if yr > TODAY_YR: continue
                lb = yr_label(cells[0]) or str(yr)
                items.append((yr, lb, val))
            items.sort(key=lambda x: x[0])
            return [(lb, v) for _, lb, v in items]
        return []
    except: return []

def fetch_bs_cash(code, ref_years=None):
    try:
        r = requests.get(f'https://irbank.net/{code}/bs', headers=HEADERS, timeout=15)
        soup = BeautifulSoup(r.content, 'html.parser')
        tables = soup.find_all('table')
        if not tables: return None
        t = tables[0]
        rows = t.find_all('tr')
        hdr = [c.get_text(strip=True) for c in rows[0].find_all(['th','td'])]
        cash_i = find_col(hdr, ['現金及び預金', '現金等', '現金・預金', '現金及び現金同等物'])
        if cash_i is None: return None
        cash_data = {}
        for row in rows[1:]:
            cells = [c.get_text(strip=True) for c in row.find_all(['th','td'])]
            if not cells or not cells[0]: continue
            yr_m = re.search(r'(\d{4})', cells[0])
            if not yr_m or int(yr_m.group(1)) > TODAY_YR: continue
            lb = yr_label(cells[0]) or yr_m.group(1)
            val = parse_num(cells[cash_i]) if cash_i < len(cells) else None
            if val is not None: cash_data[lb] = val
        if ref_years:
            cy = [y for y in sorted(cash_data.keys()) if y in set(ref_years)][-10:]
        else:
            cy = sorted(cash_data.keys())[-10:]
        return {'cash_years': cy, 'cash': [cash_data[y] for y in cy]}
    except: return None

# 消えた39社
missing = ['8766', '8630', '8725']

with open('C:/Claude/高配当株スクリーニング/chart_data_cache.json', encoding='utf-8') as f:
    data = json.load(f)

recovered = 0
for i, code in enumerate(missing, 1):
    print(f'[{i}/{len(missing)}] {code}', end=' ', flush=True)
    res = fetch_results(code); time.sleep(0.6)
    if not res or not res.get('years'):
        print('→ データなし')
        continue
    bs_cash = fetch_bs_cash(code, ref_years=res.get('years',[])); time.sleep(0.6)
    div = fetch_dividend(code); time.sleep(0.6)

    eps_dict = dict(zip(res['years'], res['eps']))
    div_years = [d[0] for d in div]
    div_vals  = [d[1] for d in div]
    payout_years, payout_vals = [], []
    for lb, dv in zip(div_years, div_vals):
        e = eps_dict.get(lb)
        if e is not None and e > 0 and dv is not None:
            payout_years.append(lb)
            payout_vals.append(round(dv / e * 100, 1))

    if bs_cash and bs_cash.get('cash'):
        cash_years = bs_cash['cash_years']
        cash = bs_cash['cash']
    elif res.get('cf_cash'):
        cash_years = res.get('cf_cash_years', [])
        cash = res.get('cf_cash', [])
    else:
        cash_years = []
        cash = []
    existing = data.get(code, {})

    entry = {
        'name': existing.get('name', code),
        'code': code,
        'yield': existing.get('yield', 0),
        'score': existing.get('score', 0),
        'fetch_date': str(datetime.date.today()),
        'years': res['years'], 'sales': res['sales'],
        'op_margin': res['op_margin'], 'eps': res['eps'],
        'cf_years': res['cf_years'], 'ocf': res['ocf'],
        'cash_years': cash_years, 'cash': cash,
        'eq_years': res['eq_years'], 'equity': res['equity'],
        'div_years': div_years, 'div': div_vals,
        'payout_years': payout_years, 'payout': payout_vals,
        'verdicts': existing.get('verdicts', {}),
    }
    has_sales = any(v is not None for v in res['sales'])
    data[code] = entry
    recovered += 1
    print(f'→ 復元 {"(収益あり)" if has_sales else "(収益なし)"}')

with open('C:/Claude/高配当株スクリーニング/chart_data_cache.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False)
print(f'\n復元完了: {recovered}件 / 合計: {len(data)}銘柄')
