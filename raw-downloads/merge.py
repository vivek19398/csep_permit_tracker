import pandas as pd, re, glob, os

OUT = '/Users/atlasdrifter/Permit'
MONTHS = {'jan':'Jan','feb':'Feb','mar':'Mar','apr':'Apr','may':'May','jun':'Jun','jul':'Jul','aug':'Aug','sep':'Sep','oct':'Oct','nov':'Nov','dec':'Dec',
          'january':'Jan','february':'Feb','march':'Mar','april':'Apr','june':'Jun','july':'Jul','august':'Aug','september':'Sep','october':'Oct','november':'Nov','december':'Dec'}

def s(v): return str(v).strip() if pd.notna(v) else ''
def year_of(f): return int(re.search(r'(20\d\d)', f).group(1))
def raw(f): return pd.read_excel(f, header=None)

def find_row(df, *terms):
    for i, row in df.iterrows():
        cells = [s(c).lower() for c in row]
        if all(any(t in c for c in cells) for t in terms):
            return i
    return None

def colmap(row):
    m = {}
    for j, c in enumerate(row):
        c = s(c).lower()
        if c == 'new': m['New'] = j
        elif c.startswith('renewal'): m['Renewal'] = j
        elif c in ('total', 'issued'): m['Issued'] = j
        elif c == 'refused': m['Refused'] = j
        elif c == 'withdrawn': m['Withdrawn'] = j
    return m

def label_table(f, label_term, label_name):
    """Nationality / county style tables -> long rows."""
    df = raw(f); y = year_of(f)
    h = find_row(df, 'refused')
    row = df.iloc[h]
    lab = next((j for j, c in enumerate(row) if label_term in s(c).lower()), 0)
    cm = colmap(row)
    out = []
    for i in range(h + 1, len(df)):
        name = s(df.iat[i, lab])
        if not name or name.lower() in ('grand total', 'total', 'year', 'jan - dec'): continue
        rec = {'Year': y, label_name: name}
        for k, j in cm.items(): rec[k] = df.iat[i, j]
        out.append(rec)
    return out

def sector_rows(f):
    df = raw(f); y = year_of(f)
    h = find_row(df, 'refused')
    out = []
    if h is not None and 'month' in [s(c).lower() for c in df.iloc[h]]:
        # old nested format: Year, Month, Sector, New, Renewal, Total, Refused, Withdrawn
        row = df.iloc[h]
        mcol = [s(c).lower() for c in row].index('month')
        scol = mcol + 1
        if s(row[scol]).lower() not in ('', 'sector'): scol = [s(c).lower() for c in row].index('sector')
        cm = colmap(row)
        cur = None
        for i in range(h + 1, len(df)):
            mv = s(df.iat[i, mcol]).lower()
            if mv in MONTHS:
                cur = MONTHS[mv]
                continue  # month subtotal row, even if a sector label bled onto it
            sec = s(df.iat[i, scol])
            if not sec or sec.lower() in ('total', 'grand total') or cur is None: continue
            rec = {'Year': y, 'Month': cur, 'Sector': sec}
            for k, j in cm.items(): rec[k] = df.iat[i, j]
            out.append(rec)
        return out
    # wide format: months as columns
    mrow = None
    for i, row in df.iterrows():
        cols = {j: MONTHS[s(c).lower()] for j, c in enumerate(row) if s(c).lower() in MONTHS}
        if len(cols) >= 3: mrow, mcols = i, cols; break
    for i in range(mrow + 1, len(df)):
        sec = s(df.iat[i, 0])
        if not sec or sec.lower() in ('economic sector', 'grand total', 'total'): continue
        for j, mon in mcols.items():
            v = df.iat[i, j]
            if pd.notna(v): out.append({'Year': y, 'Month': mon, 'Sector': sec, 'Issued': v})
    return out

def company_rows(f):
    df = raw(f); y = year_of(f)
    h = find_row(df, 'employer name')
    if h is None:  # 2025 format: no header label, months + Grand Total columns
        tcol = next(j for j, c in enumerate(df.iloc[0]) if 'grand total' in s(c).lower())
        out = []
        for i in range(1, len(df)):
            name = s(df.iat[i, 0])
            if not name or name.lower() in ('grand total', 'total'): continue
            v = pd.to_numeric(df.iat[i, tcol], errors='coerce')
            if pd.notna(v): out.append({'Year': y, 'Employer Name': name, 'Total': int(v)})
        return out
    row = df.iloc[h]
    cells = [s(c).lower() for c in row]
    ecol = next(j for j, c in enumerate(cells) if 'employer name' in c)
    if 'month' in cells:  # nested 2018/2019 format
        tcol = cells.index('total')
        recs = {}
        for i in range(h + 1, len(df)):
            name = s(df.iat[i, ecol])
            if not name or name.lower() in ('jan - dec', 'grand total'): continue
            v = pd.to_numeric(df.iat[i, tcol], errors='coerce')
            if pd.notna(v): recs[name] = recs.get(name, 0) + v
        return [{'Year': y, 'Employer Name': n, 'Total': int(t)} for n, t in recs.items()]
    # total column: 'grand total' in header rows <= h, else 'total'
    tcol = None
    for i in range(h + 1):
        for j, c in enumerate(df.iloc[i]):
            if 'grand total' in s(c).lower(): tcol = j
    if tcol is None: tcol = cells.index('total')
    out = []
    for i in range(h + 1, len(df)):
        name = s(df.iat[i, ecol])
        if not name or name.lower() in ('grand total', 'total'): continue
        v = pd.to_numeric(df.iat[i, tcol], errors='coerce')
        if pd.notna(v): out.append({'Year': y, 'Employer Name': name, 'Total': int(v)})
    return out

def build(patterns, fn, cols, outname):
    files = sorted([f for p in patterns for f in glob.glob(p)], key=year_of)
    rows = []
    for f in files: rows += fn(f)
    df = pd.DataFrame(rows)
    for c in cols:
        if c not in df: df[c] = pd.NA
    df = df[cols]
    for c in [c for c in cols if c in ('New','Renewal','Issued','Refused','Withdrawn','Total')]:
        df[c] = pd.to_numeric(df[c], errors='coerce').astype('Int64')
    df.to_csv(os.path.join(OUT, outname), index=False)
    print(outname, len(df), 'rows from', len(files), 'files, years', df.Year.min(), '-', df.Year.max())

os.makedirs(OUT, exist_ok=True)
build(['*nationality*'], lambda f: label_table(f, 'nationality', 'Nationality'),
      ['Year','Nationality','New','Renewal','Issued','Refused','Withdrawn'], 'permits-by-nationality-2010-2026.csv')
build(['*county*'], lambda f: label_table(f, 'county', 'County'),
      ['Year','County','New','Renewal','Issued','Refused','Withdrawn'], 'permits-by-county-2010-2026.csv')
build(['*sector*'], sector_rows,
      ['Year','Month','Sector','New','Renewal','Issued','Refused','Withdrawn'], 'permits-by-sector-2010-2026.csv')
build(['*compan*'], company_rows,
      ['Year','Employer Name','Total'], 'permits-issued-to-companies-2010-2026.csv')
