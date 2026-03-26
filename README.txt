MT5 Multi-Account P/L Dashboard
=================================

FILES
-----
  run_reports.bat     - Exports HTML reports from all 10 MT5 terminals, then calls merge_summary.vbs
  merge_summary.vbs   - Parses exported reports, aggregates P/L data, writes dashboard.html
  dashboard.html      - Output dashboard (overwritten each run); open in any browser


FOLDER LAYOUT ON THE WINDOWS SERVER
-------------------------------------
  C:\MT5_Accounts\
      ACC_001\terminal64.exe   (+ portable.ini)
      ACC_002\terminal64.exe
      ...
      ACC_010\terminal64.exe

  C:\MT5_Reports\
      run_reports.bat
      merge_summary.vbs
      dashboard.html           (auto-generated output)
      reports\
          ACC_001.html         (MT5-exported reports — auto-generated)
          ACC_002.html
          ...
          ACC_010.html


FIRST-TIME SETUP
-----------------
1. Copy all files into C:\MT5_Reports\
2. Open run_reports.bat in Notepad and fill in real account credentials:
     SET ACC_001_LOGIN=<your login>
     SET ACC_001_PASS=<your password>
     SET ACC_001_SERVER=<broker server name>
   Repeat for all 10 accounts.
3. Make sure each MT5 terminal folder contains a valid portable.ini.
   Create a minimal portable.ini if missing:
     [Common]
     PortablePath=1
4. Double-click run_reports.bat to do a test run.
5. Open C:\MT5_Reports\dashboard.html in your browser to verify data.


TASK SCHEDULER — HOURLY UPDATES
---------------------------------
1. Open Task Scheduler  (Win+R → taskschd.msc → Enter)
2. Click "Create Basic Task" in the right panel.
3. Name:    MT5 PnL Dashboard
   Description: Exports MT5 reports and regenerates P/L dashboard every hour
4. Trigger: Daily  →  set start time (e.g. 07:00)
5. Action:  Start a program
   Program/script:  C:\MT5_Reports\run_reports.bat
6. Finish the wizard, then open the task Properties:
   Triggers tab → Edit trigger:
     [x] Repeat task every:  1 hour
         for a duration of:  Indefinitely
7. Settings tab:
     [x] Run task as soon as possible after a scheduled start is missed
     [x] If the task fails, restart every: 5 minutes, up to 3 times
8. General tab:
     [x] Run whether user is logged on or not
     [x] Run with highest privileges  (if C:\MT5_Accounts needs admin rights)
9. Click OK, enter your Windows password when prompted.


VIEWING THE DASHBOARD
----------------------
Open C:\MT5_Reports\dashboard.html in any browser.
The page auto-refreshes every 5 minutes so you can leave it open on a monitor.
For remote access, share the file via a local network share or a web server (IIS, nginx).


TROUBLESHOOTING
----------------
- Report file not created after run:
    Check that terminal64.exe path is correct in run_reports.bat.
    Increase WAIT_SECONDS if MT5 needs more time to connect (slow broker).
    Verify the broker server name matches exactly what MT5 Manager shows.

- All values show 0.00 in dashboard:
    The HTML parser in merge_summary.vbs looks for MT5's standard label text
    ("Balance", "Equity", "Profit for today", etc.).
    Open one of the exported acc_00X.html files and check the exact label wording.
    Update ParseMetric / ParsePLPeriod in merge_summary.vbs to match.

- MT5 window flashes briefly and report is empty:
    MT5 connected but could not authenticate.  Double-check login/password/server.

- Script blocked by execution policy (if run via PowerShell):
    Run cscript directly:  cscript //NoLogo C:\MT5_Reports\merge_summary.vbs
    The .bat file already uses cscript so this should not be an issue.
