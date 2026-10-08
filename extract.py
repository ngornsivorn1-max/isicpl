#!/usr/bin/env python3
"""ISI E&C management review — data extractor.

Reads the three source workbooks and writes the JSON files the review site loads:
  pnl.json  <- Master_Project_CashFlow_Dashboard_CPL_MASTER.xlsm (Exec Dashboard, Summary, Project Master, ISI CON, ISI BDS)
  con.json  <- Master_CPL_CON_Workbook (Dashboard, List View)
  bds.json  <- Master_CPL_BDS_Workbook (Dashboard, List View)

Usage:  python3 extract.py <cashflow.xlsm> <cpl_con.xlsx> <cpl_bds.xlsx> <out_dir>
Every number is read from cached cell values (data_only) — Excel must have been saved after the last recalc.
"""
import sys, json, datetime, collections, re, os, warnings
warnings.filterwarnings('ignore')
import openpyxl

MONTHS = [f"{y}-{m:02d}" for y in (2025, 2026, 2027) for m in range(1, 13)]
def ym(d): return d.strftime('%Y-%m') if isinstance(d, datetime.datetime) else None
def dstr(d): return d.strftime('%Y-%m-%d') if isinstance(d, datetime.datetime) else None
def num(v): return round(float(v), 2) if isinstance(v, (int, float)) else 0.0
def s(v): return str(v).strip() if v is not None else ''
def r2(x): return round(x, 2)

# ----------------------------------------------------------------------------- P&L (cash-flow master)
def extract_pnl(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ex = wb['Exec Dashboard']
    as_of = s(ex['M1'].value).replace('DATA AS OF:', '').strip()
    refresh = s(ex['M2'].value).replace('Last Refresh:', '').strip()
    pm = wb['Project Master']
    gp_target = pm['B4'].value or 0.15
    coll_target = pm['D4'].value or 0.9

    # project table: header row has 'Project Code' in col A
    hdr_row = None
    for i, row in enumerate(pm.iter_rows(min_row=40, max_row=60, max_col=1, values_only=True), 40):
        if row[0] == 'Project Code': hdr_row = i
    # columns located by header name so inserted / reordered columns in the master do not shift the figures
    import re as _re
    hdr = next(pm.iter_rows(min_row=hdr_row, max_row=hdr_row, values_only=True))
    norm = lambda v: _re.sub(r'[^a-z0-9$%]', '', str(v or '').lower())
    hmap = {}
    for i, h in enumerate(hdr):
        k = norm(h)
        if k and k not in hmap: hmap[k] = i
    def col(*aliases):
        for a in aliases:
            i = hmap.get(norm(a))
            if i is not None: return i
        return -1
    C = dict(code=col('Project Code'), bu=col('BU'), name=col('Project Name'), start=col('Actual Start Date', 'Start Date', 'Start'), finish=col('Estimate to Finish Date', 'Finish Date', 'Finish'), handover=col('Handover Date', 'Handover'),
        status=col('Status', 'Project Status'), contract=col('Contract Value $', 'Contract $', 'Current Value $'), fcstRev=col('Fcst Revenue $', 'Forecast Revenue $'), actRev=col('Actual Revenue $'),
        apprCost=col('Approved Cost $'), revCost=col('Revise Cost $', 'Revised Cost $'), actCost=col('Actual Cost $'), apprGP=col('Approved GP $'), revGP=col('Revise GP $', 'Revised GP $'), actGP=col('Actual GP $'),
        eac=col('EAC Cost $', 'EAC $', 'Estimate at Completion $'), etc=col('ETC Cost $', 'ETC $', 'Estimate to Complete $'),
        apprGPp=col('Approved GP %'), revGPp=col('Revise GP %', 'Revised GP %'), actGPp=col('Actual GP %'), eacGPp=col('EAC GP %', 'ETC GP %'),
        invoiced=col('Invoiced (incl VAT) $', 'Invoiced $'), collected=col('Cash Collected $', 'Collected $'), collPct=col('Collection %'), ar=col('Actual AR $', 'Outstanding AR $', 'AR $'),
        future=col('Future Cash to Collect $', 'Future Cash $'), gpVar=col('GP Var $', 'GP Variance $'), gpVarPts=col('GP Var pts', 'GP Var %'), costVar=col('Cost Var $', 'Cost Variance $'),
        health=col('Health'), risk=col('Main Risk', 'Risk'), riskScore=col('Risk Score $', 'Risk Score'), family=col('Family'))
    missing = [k for k, v in C.items() if v < 0 and k not in ('eac', 'etc', 'eacGPp', 'handover')]
    if missing: raise SystemExit(f'Project Master: columns not found by header name: {missing}')
    if C['eac'] < 0 and C['etc'] < 0: raise SystemExit('Project Master: neither "EAC Cost $" nor "ETC Cost $" found')
    g = lambda r, i: (r[i] if 0 <= i < len(r) else None)
    pct = lambda v: v if isinstance(v, (int, float)) else None
    rows = [r for r in pm.iter_rows(min_row=hdr_row + 1, max_col=len(hdr), values_only=True) if g(r, C['code'])]
    projects = []
    for r in rows:
        nm, st = g(r, C['name']), g(r, C['status'])
        projects.append(dict(
            code=s(g(r, C['code'])), bu=s(g(r, C['bu'])), name=s(nm) if nm not in (None, 0) else '', start=dstr(g(r, C['start'])), finish=dstr(g(r, C['finish'])), handover=dstr(g(r, C['handover'])),
            status=s(st) if st not in (None, 0) else 'Not set', contract=num(g(r, C['contract'])), fcstRev=num(g(r, C['fcstRev'])), actRev=num(g(r, C['actRev'])),
            apprCost=num(g(r, C['apprCost'])), revCost=num(g(r, C['revCost'])), actCost=num(g(r, C['actCost'])), apprGP=num(g(r, C['apprGP'])), revGP=num(g(r, C['revGP'])), actGP=num(g(r, C['actGP'])),
            etc=num(g(r, C['eac'] if C['eac'] >= 0 else C['etc'])), apprGPp=pct(g(r, C['apprGPp'])), revGPp=pct(g(r, C['revGPp'])), actGPp=pct(g(r, C['actGPp'])), etcGPp=pct(g(r, C['eacGPp'])),
            invoiced=num(g(r, C['invoiced'])), collected=num(g(r, C['collected'])), collPct=pct(g(r, C['collPct'])), ar=num(g(r, C['ar'])), future=num(g(r, C['future'])), gpVar=num(g(r, C['gpVar'])),
            gpVarPts=pct(g(r, C['gpVarPts'])), costVar=num(g(r, C['costVar'])), health=s(g(r, C['health'])), risk=s(g(r, C['risk'])), riskScore=num(g(r, C['riskScore'])), family=s(g(r, C['family']))))
    # an "ETC Cost $" column may hold the estimate AT completion (older masters: equals revise cost) or the estimate TO complete (EAC − actual); the app stores EAC
    if C['eac'] < 0 and projects:
        eq = sum(1 for p in projects if abs(p['etc'] - p['revCost']) < 1)
        if eq / len(projects) < 0.5:
            for p in projects: p['etc'] = p['etc'] + p['actCost']
    # customer + monthly forecast cash-in + AR aging per project from ISI CON / ISI BDS registers
    reg = {}
    profiles = {}
    for sheet, bu in (('ISI CON', 'ISI CON'), ('ISI BDS', 'ISI BDS')):
        ws = wb[sheet]
        hdr = None; prof_hdr = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=120, max_col=2, values_only=True), 1):
            if row[0] == 'Project Code' and hdr is None: hdr = i
            if row[0] == 'Month' and row[1] == 'Planning $': prof_hdr = i
        h = list(next(ws.iter_rows(min_row=hdr, max_row=hdr, values_only=True)))
        mcols = [(j, ym(v)) for j, v in enumerate(h) if isinstance(v, datetime.datetime)]
        ar_cols = {k: h.index(k) for k in ('AR Not due', 'AR 1-30', 'AR 31-60', 'AR 61-90', 'AR >90')}
        for r in ws.iter_rows(min_row=hdr + 1, values_only=True):
            if not r[0]: continue
            code = s(r[0])
            reg[code] = dict(customer=s(r[2]), fcastIn={m: num(r[j]) for j, m in mcols if num(r[j])}, post2027=num(r[h.index('Post-2027')]),
                             aging=[num(r[ar_cols[k]]) for k in ('AR Not due', 'AR 1-30', 'AR 31-60', 'AR 61-90', 'AR >90')])
        # KPI row 7, by-status table, aging table, family table, collection profile (planning vs actual cash in + expenses)
        k = list(next(ws.iter_rows(min_row=7, max_row=7, max_col=14, values_only=True)))
        kpi = dict(projects=k[0], contractOrig=num(k[1]), vo=num(k[2]), contract=num(k[3]), cost=num(k[4]), margin=num(k[5]), marginPct=k[6],
                   progress=k[7], invoiced=num(k[8]), collected=num(k[9]), remaining=num(k[10]), overdueFinished=num(k[11]), fcastIn=num(k[12]), fcastOut=num(k[13]))
        by_status = []
        for r in ws.iter_rows(min_row=11, max_row=24, max_col=4, values_only=True):
            if r[0] in (None, 0) or not isinstance(r[1], (int, float)): continue
            by_status.append(dict(status=s(r[0]), lines=r[1], value=num(r[2]), remaining=num(r[3])))
        ar_status = []
        st = None
        for i, r in enumerate(ws.iter_rows(min_row=25, max_row=45, max_col=4, values_only=True), 25):
            if r[0] == 'A/R BY PROJECT STATUS': st = i + 2; continue
            if st and i >= st:
                if r[0] in (None,) or r[0] == 'COLLECTION PROFILE Pre-2025 → Post-2027 ($) — PLANNING vs ACTUAL': break
                if r[0] != 0 and isinstance(r[1], (int, float)): ar_status.append(dict(status=s(r[0]), invoiced=num(r[1]), collected=num(r[2]), remaining=num(r[3])))
        aging = []
        for r in ws.iter_rows(min_row=27, max_row=32, min_col=6, max_col=8, values_only=True):
            if r[0] and r[0] != 'Total': aging.append(dict(bucket=s(r[0]), amount=num(r[1]), lines=r[2] or 0))
        family = []
        for r in ws.iter_rows(min_row=7, max_row=30, min_col=18, max_col=26, values_only=True):
            if r[0] in (None, 'TOTAL') or not isinstance(r[1], (int, float)): continue
            family.append(dict(family=s(r[0]), projects=r[1], value=num(r[2]), cost=num(r[3]), margin=num(r[4]), marginPct=r[5], invoiced=num(r[6]), collected=num(r[7]), ar=num(r[8])))
        profile = []
        for r in ws.iter_rows(min_row=prof_hdr + 1, max_row=prof_hdr + 40, max_col=10, values_only=True):
            if r[0] is None or (isinstance(r[0], str) and r[0].startswith('Total')): break
            profile.append(dict(month=ym(r[0]) or s(r[0]), planIn=num(r[1]), actIn=num(r[3]), planOut=num(r[5]), actOut=num(r[7]), etcOut=num(r[9])))
        profiles[bu] = dict(kpi=kpi, byStatus=by_status, arByStatus=ar_status, aging=aging, family=family, profile=profile)
    for p in projects:
        g = reg.get(p['code'], {})
        p['customer'] = g.get('customer', ''); p['fcastIn'] = g.get('fcastIn', {}); p['post2027'] = g.get('post2027', 0); p['aging'] = g.get('aging', [0, 0, 0, 0, 0])
    # per-project monthly profile from the MG lines (Planning / Actual / Forecast rows × monthly $ block) — what the sheet's collection profile sums
    MON = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
    def mkey(v):
        if isinstance(v, datetime.datetime): return ym(v)
        t = s(v)
        if t in ('Pre-2025', 'Post-2027'): return t
        m = re.match(r'^([A-Za-z]{3})-(\d{2})$', t)
        return f'20{m.group(2)}-{MON[m.group(1)]}' if m and m.group(1) in MON else None
    monthly = {}
    for sheet in ('MG CON Lines', 'MG BDS Lines'):
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        hdr_i = None
        for i, r in enumerate(ws.iter_rows(min_row=1, max_row=15, min_col=5, max_col=5, values_only=True), 1):
            if s(r[0]).startswith('Project Code'): hdr_i = i; break
        if not hdr_i: continue
        hr = list(ws.iter_rows(min_row=hdr_i, max_row=hdr_i, values_only=True))[0]
        pres = [j for j, v in enumerate(hr) if s(v) == 'Pre-2025']
        start = pres[1] if len(pres) >= 2 else (pres[0] if pres else None)
        if start is None: continue
        cols = []
        for j in range(start, len(hr)):
            k = mkey(hr[j])
            if not k: break
            cols.append((j, k))
            if k == 'Post-2027': break
        for r in ws.iter_rows(min_row=hdr_i + 1, max_row=ws.max_row, values_only=True):
            code, typ = s(r[4]) if len(r) > 4 else '', s(r[1]) if len(r) > 1 else ''
            if not code or typ not in ('Planning', 'Actual', 'Forecast'): continue
            m = monthly.setdefault(code, {'plan': {}, 'act': {}, 'fcst': {}})
            tgt = m['plan'] if typ == 'Planning' else m['act'] if typ == 'Actual' else m['fcst']
            for j, k in cols:
                v = num(r[j]) if j < len(r) else 0
                if v: tgt[k] = round(tgt.get(k, 0) + v, 2)
    for p in projects:
        m = monthly.get(p['code'])
        if m: p['mPlan'] = m['plan']; p['mAct'] = m['act']; p['mFcst'] = m['fcst']
    # Exec Dashboard blocks
    health = []
    for r in ex.iter_rows(min_row=19, max_row=22, min_col=10, max_col=14, values_only=True):
        if r[0]: health.append(dict(status=s(r[0]), projects=r[1], contract=num(r[2]), actRev=num(r[3]), actGP=num(r[4])))
    dist = []
    for r in ex.iter_rows(min_row=46, max_row=51, min_col=10, max_col=12, values_only=True):
        if r[0]: dist.append(dict(range=s(r[0]), projects=r[1], contract=num(r[2])))
    alerts = [s(r[0]) for r in ex.iter_rows(min_row=18, max_row=24, max_col=1, values_only=True) if r[0]]
    controls = []
    for r in ex.iter_rows(min_row=82, max_row=90, max_col=3, values_only=True):
        if r[0]: controls.append(dict(check=s(r[0]), result=r[1] if isinstance(r[1], (int, float)) else s(r[1]), status=s(r[2])))
    fc_monthly = []
    for r in ex.iter_rows(min_row=21, max_row=32, min_col=23, max_col=26, values_only=True):
        if isinstance(r[0], datetime.datetime): fc_monthly.append(dict(month=ym(r[0]), con=num(r[1]), bds=num(r[2]), total=num(r[3])))
    perf = []
    for r in ex.iter_rows(min_row=11, max_row=14, max_col=6, values_only=True):
        perf.append(dict(kpi=s(r[0]), forecast=num(r[1]), actual=num(r[2]), variance=num(r[3]), varPct=r[4] if isinstance(r[4], (int, float)) else None, assessment=s(r[5])))
    notes = [s(r[0]) for r in ex.iter_rows(min_row=91, max_row=91, max_col=1, values_only=True) if r[0]]
    legend = []
    for r in pm.iter_rows(min_row=1, max_row=4, min_col=12, max_col=24, values_only=True):
        rw = [str(v).strip() for v in r if isinstance(v, str) and v.strip()]
        if len(rw) > 1: legend.append(rw)
    vat = wb['Summary']['E30'].value or 0.1
    return dict(asOf=as_of, refresh=refresh, gpTarget=gp_target, collTarget=coll_target, vat=vat, projects=projects, bu=profiles,
                health=health, distribution=dist, alerts=alerts, controls=controls, fcMonthly=fc_monthly, perf=perf, notes=notes, legend=legend)

# ----------------------------------------------------------------------------- CPL workbooks (Dashboard + List View)
def extract_cpl(path, division):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Dashboard']
    title = s(ws['B1'].value)
    tot = list(next(ws.iter_rows(min_row=6, max_row=6, min_col=2, max_col=17, values_only=True)))
    totals = dict(budget=num(tot[0]), budgetAwarded=num(tot[1]), awarded=num(tot[2]), saving=num(tot[3]), savingPct=tot[4], items=tot[5] or 0, itemsAwarded=tot[6] or 0,
                  within=tot[7] or 0, over=tot[8] or 0, late=tot[9] or 0, plannedPct=tot[10], actualPct=tot[11], remainingPct=tot[12], scheduled=num(tot[13]), paid=num(tot[14]), payStatus=s(tot[15]))
    projects = []; section = 'main'
    late_steps = None; reasons = []; cash = []
    rows = list(ws.iter_rows(min_row=9, max_row=ws.max_row, max_col=18, values_only=True))
    i = 0
    while i < len(rows):
        r = rows[i]
        if r[1] == 'ADDED PROJECTS (from claims)': section = 'added'
        if isinstance(r[0], int) and r[1] is not None and r[2] is not None and r[2] != 'UNPOSTED CLAIMS (not yet in P&F)':
            projects.append(dict(key=f"{s(r[1])}|{s(r[2])}", code=s(r[1]), name=s(r[2]), section=section, items=r[3] or 0, awardedItems=r[4] or 0, budget=num(r[5]), budgetAwarded=num(r[6]),
                                 awarded=num(r[7]), saving=num(r[8]), savingPct=r[9] if isinstance(r[9], (int, float)) else None, within=r[10] or 0, over=r[11] or 0, late=r[12] or 0,
                                 scheduled=num(r[13]), paid=num(r[14]), actualPct=r[15] if isinstance(r[15], (int, float)) else None, payStatus=s(r[16]) or None, projStatus=s(r[17]) or 'Unassigned',
                                 plan={m: 0.0 for m in MONTHS}, planPost=0.0, actual={m: 0.0 for m in MONTHS}))
        if r[1] == 'Step name':
            names = [s(v) for v in r[2:16]]; cnt = rows[i + 1][2:16]
            late_steps = [dict(step=n, late=c or 0) for n, c in zip(names, cnt)]
        if r[1] == 'Reason':
            j = i + 1
            while j < len(rows) and rows[j][1] and rows[j][1] != 'TOTAL flagged':
                reasons.append([s(rows[j][1]), rows[j][2] or 0]); j += 1
        if r[1] == 'Month' and r[2] == 'Cash-Out $':
            j = i + 1
            while j < len(rows) and rows[j][1] is not None and rows[j][1] != 'TOTAL':
                cash.append(dict(month=ym(rows[j][1]) or s(rows[j][1]), plan=num(rows[j][2]), paid=num(rows[j][4]))); j += 1
        i += 1
    unposted = 0.0
    for r in rows:
        if r[2] == 'UNPOSTED CLAIMS (not yet in P&F)': unposted = num(r[14])
    pidx = {p['key']: p for p in projects}
    # List View -> items + monthly schedule per project
    lv = wb['List View']
    h = list(next(lv.iter_rows(min_row=3, max_row=3, values_only=True)))
    mcols = [(j, ym(v)) for j, v in enumerate(h) if isinstance(v, datetime.datetime)]
    post = h.index('Post-2027')
    items = []; unmatched = collections.Counter()
    for r in lv.iter_rows(min_row=4, values_only=True):
        if r[0] is None or r[3] is None or (isinstance(r[0], str) and r[0].startswith(' ')) or r[2] is None: continue
        key = f"{s(r[0])}|{s(r[1])}"
        p = pidx.get(key) or pidx.get(f"{s(r[0])}|{s(r[0])}")
        if p is None:
            # fall back: first project with same code
            cands = [q for q in projects if q['code'] == s(r[0])]
            p = cands[0] if cands else None
        if p is None: unmatched[key] += 1; continue
        for j, m in mcols:
            if isinstance(r[j], (int, float)) and m in p['plan']: p['plan'][m] += float(r[j])
        if isinstance(r[post], (int, float)): p['planPost'] += float(r[post])
        items.append(dict(k=p['key'], cat=s(r[2]), item=r[3], desc=s(r[4]), type=s(r[5]) or None, sup=s(r[6]) or None, status=s(r[7]), budget=num(r[8]), bAwd=num(r[9]),
                          awd=num(r[10]), var=num(r[11]), award=ym(r[12]), start=ym(r[13]), finish=ym(r[14]), overall=s(r[15]), paid=num(r[17]), due=num(r[18])))
    # actual claims by project by month
    ac = wb['Actual Cash-Out (Claims)']
    h2 = list(next(ac.iter_rows(min_row=4, max_row=4, values_only=True)))
    mc2 = [(j, ym(v)) for j, v in enumerate(h2) if isinstance(v, datetime.datetime)]
    for r in ac.iter_rows(min_row=5, values_only=True):
        if not r[0] or not isinstance(r[0], str) or r[0].startswith(' ') or r[3] is None: continue
        key = f"{s(r[0])}|{s(r[1])}"
        p = pidx.get(key) or pidx.get(f"{s(r[0])}|{s(r[0])}")
        if p is None:
            cands = [q for q in projects if q['code'] == s(r[0])]; p = cands[0] if cands else None
        if p is None: continue
        for j, m in mc2:
            if isinstance(r[j], (int, float)) and m in p['actual']: p['actual'][m] += float(r[j])
    for p in projects:
        p['plan'] = [r2(p['plan'][m]) for m in MONTHS]; p['actual'] = [r2(p['actual'][m]) for m in MONTHS]; p['planPost'] = r2(p['planPost'])
    return dict(division=division, title=title, totals=totals, projects=projects, items=items, lateSteps=late_steps or [], reasons=reasons, cash=cash, unposted=unposted,
                unmatchedItems=sum(unmatched.values()), months=MONTHS)

if __name__ == '__main__':
    cf, con, bds, out = sys.argv[1:5]
    os.makedirs(out, exist_ok=True)
    generated = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    pnl = extract_pnl(cf); pnl['generated'] = generated
    json.dump(pnl, open(os.path.join(out, 'pnl.json'), 'w'))
    for path, div in ((con, 'CON'), (bds, 'BDS')):
        d = extract_cpl(path, div); d['generated'] = generated
        json.dump(d, open(os.path.join(out, f'{div.lower()}.json'), 'w'))
        print(div, 'projects', len(d['projects']), 'items', len(d['items']), 'awarded', d['totals']['awarded'], 'paid', d['totals']['paid'], 'unmatched', d['unmatchedItems'])
    print('PNL projects', len(pnl['projects']), 'contract', sum(p['contract'] for p in pnl['projects']), 'asOf', pnl['asOf'])
