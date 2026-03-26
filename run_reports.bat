@echo off
REM ============================================================
REM  MT5 Multi-Account Report Exporter
REM  Run this bat to export P/L HTML reports from all 10 accounts
REM ============================================================
REM
REM  TASK SCHEDULER SETUP INSTRUCTIONS:
REM  -----------------------------------
REM  1. Open Task Scheduler (taskschd.msc)
REM  2. Click "Create Basic Task"
REM  3. Name: "MT5 PnL Dashboard Update"
REM  4. Trigger: Daily, repeat every 1 hour for 1 day (or set custom interval)
REM  5. Action: Start a program
REM     Program: C:\MT5_Reports\run_reports.bat
REM  6. For hourly repeat:
REM     - After creating, open task Properties -> Triggers -> Edit
REM     - Check "Repeat task every: 1 hour" for a duration of "Indefinitely"
REM  7. Run Whether User Is Logged On Or Not
REM  8. Run With Highest Privileges (if needed for file access)
REM
REM  After run_reports.bat finishes, it automatically calls merge_summary.vbs
REM ============================================================

REM ============================================================
REM  ACCOUNT CONFIGURATION — edit these before first run
REM ============================================================

SET MT5_BASE=C:\MT5_Accounts
SET REPORT_DIR=C:\MT5_Reports\reports
SET VBS_SCRIPT=C:\MT5_Reports\merge_summary.vbs
SET WAIT_SECONDS=15

REM Account 1
SET ACC_001_LOGIN=10001
SET ACC_001_PASS=Password1
SET ACC_001_SERVER=BrokerName-Demo

REM Account 2
SET ACC_002_LOGIN=10002
SET ACC_002_PASS=Password2
SET ACC_002_SERVER=BrokerName-Demo

REM Account 3
SET ACC_003_LOGIN=10003
SET ACC_003_PASS=Password3
SET ACC_003_SERVER=BrokerName-Demo

REM Account 4
SET ACC_004_LOGIN=10004
SET ACC_004_PASS=Password4
SET ACC_004_SERVER=BrokerName-Demo

REM Account 5
SET ACC_005_LOGIN=10005
SET ACC_005_PASS=Password5
SET ACC_005_SERVER=BrokerName-Demo

REM Account 6
SET ACC_006_LOGIN=10006
SET ACC_006_PASS=Password6
SET ACC_006_SERVER=BrokerName-Demo

REM Account 7
SET ACC_007_LOGIN=10007
SET ACC_007_PASS=Password7
SET ACC_007_SERVER=BrokerName-Demo

REM Account 8
SET ACC_008_LOGIN=10008
SET ACC_008_PASS=Password8
SET ACC_008_SERVER=BrokerName-Demo

REM Account 9
SET ACC_009_LOGIN=10009
SET ACC_009_PASS=Password9
SET ACC_009_SERVER=BrokerName-Demo

REM Account 10
SET ACC_010_LOGIN=10010
SET ACC_010_PASS=Password10
SET ACC_010_SERVER=BrokerName-Demo

REM ============================================================
REM  Create output directory if not exists
REM ============================================================
IF NOT EXIST "%REPORT_DIR%" (
    mkdir "%REPORT_DIR%"
    echo [INFO] Created directory: %REPORT_DIR%
)

echo ============================================================
echo  MT5 Report Export — %DATE% %TIME%
echo ============================================================

REM ============================================================
REM  Export reports for each account
REM ============================================================

CALL :ExportReport ACC_001 %ACC_001_LOGIN% %ACC_001_PASS% %ACC_001_SERVER%
CALL :ExportReport ACC_002 %ACC_002_LOGIN% %ACC_002_PASS% %ACC_002_SERVER%
CALL :ExportReport ACC_003 %ACC_003_LOGIN% %ACC_003_PASS% %ACC_003_SERVER%
CALL :ExportReport ACC_004 %ACC_004_LOGIN% %ACC_004_PASS% %ACC_004_SERVER%
CALL :ExportReport ACC_005 %ACC_005_LOGIN% %ACC_005_PASS% %ACC_005_SERVER%
CALL :ExportReport ACC_006 %ACC_006_LOGIN% %ACC_006_PASS% %ACC_006_SERVER%
CALL :ExportReport ACC_007 %ACC_007_LOGIN% %ACC_007_PASS% %ACC_007_SERVER%
CALL :ExportReport ACC_008 %ACC_008_LOGIN% %ACC_008_PASS% %ACC_008_SERVER%
CALL :ExportReport ACC_009 %ACC_009_LOGIN% %ACC_009_PASS% %ACC_009_SERVER%
CALL :ExportReport ACC_010 %ACC_010_LOGIN% %ACC_010_PASS% %ACC_010_SERVER%

echo.
echo [INFO] All reports exported. Running dashboard generator...
cscript //NoLogo "%VBS_SCRIPT%"

echo.
echo [INFO] Dashboard updated: C:\MT5_Reports\dashboard.html
echo [DONE] %DATE% %TIME%
goto :EOF

REM ============================================================
REM  Subroutine: ExportReport <AccID> <Login> <Pass> <Server>
REM ============================================================
:ExportReport
SET ACC_ID=%1
SET LOGIN=%2
SET PASS=%3
SET SERVER=%4
SET TERMINAL=%MT5_BASE%\%ACC_ID%\terminal64.exe
SET OUTFILE=%REPORT_DIR%\%ACC_ID%.html

echo [INFO] Exporting %ACC_ID% (login: %LOGIN%, server: %SERVER%)...

IF NOT EXIST "%TERMINAL%" (
    echo [WARN] Terminal not found: %TERMINAL% — skipping
    goto :ExportDone
)

REM Start MT5, export report, then detach (auto-close)
start "" /wait "%TERMINAL%" /portable /login:%LOGIN% /password:%PASS% /server:%SERVER% /report:"%OUTFILE%" /detach

REM Wait for MT5 to connect, generate and save report
timeout /t %WAIT_SECONDS% /nobreak >nul

IF EXIST "%OUTFILE%" (
    echo [OK]   %ACC_ID% report saved: %OUTFILE%
) ELSE (
    echo [WARN] %ACC_ID% report not found after export — check credentials/server
)

:ExportDone
goto :EOF
