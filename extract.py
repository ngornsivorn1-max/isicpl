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
    rows = [r for r in pm.iter_rows(min_row=hdr_row + 1, max_col=33, values_only=True) if r[0]]
    projects = []
    for r in rows:
        projects.append(dict(
            code=s(r[0]), bu=s(r[1]), name=s(r[2]) if r[2] not in (None, 0) else '', start=dstr(r[3]), finish=dstr(r[4]), handover=dstr(r[5]),
            status=s(r[6]) if r[6] not in (None, 0) else 'Not set', contract=num(r[7]), fcstRev=num(r[8]), actRev=num(r[9]),
            apprCost=num(r[10]), revCost=num(r[11]), actCost=num(r[12]), apprGP=num(r[13]), revGP=num(r[14]), actGP=num(r[15]),
            etc=num(r[16]), apprGPp=r[17] if isinstance(r[17], (int, float)) else None, revGPp=r[18] if isinstance(r[18], (int, float)) else None,
            actGPp=r[19] if isinstance(r[19], (int, float)) else None, etcGPp=r[20] if isinstance(r[20], (int, float)) else None, invoiced=num(r[21]), collected=num(r[22]),
            collPct=r[23] if isinstance(r[23], (int, float)) else None, ar=num(r[24]), future=num(r[25]), gpVar=num(r[26]),
            gpVarPts=r[27] if isinstance(r[27], (int, float)) else None, costVar=num(r[28]), health=s(r[29]), risk=s(r[30]), riskScore=num(r[31]), family=s(r[32])))
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
