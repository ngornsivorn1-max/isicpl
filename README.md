ISI E&C - CPL - P&L and Procurement Planning  (static website)

FILES
  index.html          the whole app (one file)
  data/pnl.json       P&L data   (from Master_Project_CashFlow_Dashboard_CPL_MASTER.xlsm)
  data/con.json       CPL CON    (from Master_CPL_CON_Workbook)
  data/bds.json       CPL BDS    (from Master_CPL_BDS_Workbook)
  xlsx.full.min.js    Excel reader used by "Update from Excel"
  og.png, favicon.svg link preview picture + tab icon
  vercel.json         hosting settings
  extract.py          optional: rebuilds the three JSON files on a PC with Python (python extract.py)

HOW TO PUT IT ONLINE (same as encglcode.vercel.app)
  1. Create a GitHub repository (e.g. isicpl) and upload every file in this folder, keeping the data/ folder.
  2. On vercel.com -> Add New -> Project -> Import that repository.
     Framework preset: Other. Build command: (empty). Output directory: (empty / root). Deploy.
  3. Settings -> Domains: set the project name you want, e.g. isicpl  ->  https://isicpl.vercel.app
     (the name must be free on Vercel; the app title in the link preview is already set).

HOW TO UPDATE THE FIGURES
  Open the site -> "Update from Excel" -> drop the saved workbook(s) -> "Download data files for the website".
  Replace data/pnl.json, data/con.json, data/bds.json in the repository with the downloaded files. Vercel
  redeploys automatically in about a minute and everyone sees the new figures.
