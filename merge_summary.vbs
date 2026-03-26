' ============================================================
' MT5 Multi-Account P/L Dashboard Generator
' merge_summary.vbs
'
' Reads 10 MT5-exported HTML reports from C:\MT5_Reports\reports\
' Parses balance, equity and P/L figures, then writes dashboard.html
' ============================================================

Option Explicit

' ---- Configuration ------------------------------------------
Const REPORT_DIR    = "C:\MT5_Reports\reports\"
Const OUTPUT_FILE   = "C:\MT5_Reports\dashboard.html"
Const ACCOUNT_COUNT = 10
' -------------------------------------------------------------

Dim fso, i, accID, reportPath
Dim htmlContent

Set fso = CreateObject("Scripting.FileSystemObject")

' Arrays to hold per-account data
Dim accIDs(9), accBalance(9), accEquity(9)
Dim accPLToday(9), accPLWeek(9), accPLMonth(9), accPLYear(9)
Dim accFound(9)

' Aggregated totals
Dim totalBalance, totalEquity
Dim totalPLToday, totalPLWeek, totalPLMonth, totalPLYear
totalBalance  = 0
totalEquity   = 0
totalPLToday  = 0
totalPLWeek   = 0
totalPLMonth  = 0
totalPLYear   = 0

' ============================================================
' Parse each account report
' ============================================================
For i = 0 To ACCOUNT_COUNT - 1
    accID = "ACC_" & Right("00" & CStr(i + 1), 3)
    accIDs(i) = accID
    reportPath = REPORT_DIR & accID & ".html"

    accBalance(i)  = 0
    accEquity(i)   = 0
    accPLToday(i)  = 0
    accPLWeek(i)   = 0
    accPLMonth(i)  = 0
    accPLYear(i)   = 0
    accFound(i)    = False

    If fso.FileExists(reportPath) Then
        accFound(i) = True
        htmlContent = ReadFile(fso, reportPath)

        accBalance(i)  = ParseMetric(htmlContent, "Balance")
        accEquity(i)   = ParseMetric(htmlContent, "Equity")
        accPLToday(i)  = ParsePLPeriod(htmlContent, "today")
        accPLWeek(i)   = ParsePLPeriod(htmlContent, "week")
        accPLMonth(i)  = ParsePLPeriod(htmlContent, "month")
        accPLYear(i)   = ParsePLPeriod(htmlContent, "year")

        totalBalance  = totalBalance  + accBalance(i)
        totalEquity   = totalEquity   + accEquity(i)
        totalPLToday  = totalPLToday  + accPLToday(i)
        totalPLWeek   = totalPLWeek   + accPLWeek(i)
        totalPLMonth  = totalPLMonth  + accPLMonth(i)
        totalPLYear   = totalPLYear   + accPLYear(i)
    End If
Next

' ============================================================
' Build and write dashboard HTML
' ============================================================
Dim dash, ts
ts = Now()

dash = BuildDashboard(accIDs, accBalance, accEquity, accPLToday, accPLWeek, accPLMonth, accPLYear, accFound, _
                      totalBalance, totalEquity, totalPLToday, totalPLWeek, totalPLMonth, totalPLYear, ts)

WriteFile fso, OUTPUT_FILE, dash

WScript.Echo "Dashboard written to: " & OUTPUT_FILE

' ============================================================
' Helper: Read entire file as string
' ============================================================
Function ReadFile(objFSO, path)
    Dim f
    Set f = objFSO.OpenTextFile(path, 1)
    ReadFile = f.ReadAll()
    f.Close
End Function

' ============================================================
' Helper: Write string to file (overwrite)
' ============================================================
Sub WriteFile(objFSO, path, content)
    Dim f
    Set f = objFSO.CreateTextFile(path, True)
    f.Write content
    f.Close
End Sub

' ============================================================
' ParseMetric: extract numeric value following a label
' MT5 HTML report format uses table rows like:
'   <td ...>Balance:</td><td ...>12 345.67</td>
' ============================================================
Function ParseMetric(html, label)
    Dim pattern, pos, valueStart, valueEnd, rawVal, cleanVal
    Dim lowerHtml
    lowerHtml = LCase(html)
    label = LCase(label)

    ' Find label occurrence
    pos = InStr(lowerHtml, ">" & label & "<")
    If pos = 0 Then
        pos = InStr(lowerHtml, ">" & label & ":<")
        If pos = 0 Then
            ParseMetric = 0
            Exit Function
        End If
    End If

    ' Find the next <td> closing tag after label and grab value
    Dim afterLabel
    afterLabel = Mid(html, pos, 500)

    ' Locate opening of next cell value — skip past </td>
    Dim tdClose
    tdClose = InStr(afterLabel, "</td>")
    If tdClose = 0 Then
        ParseMetric = 0
        Exit Function
    End If
    Dim afterClose
    afterClose = Mid(afterLabel, tdClose + 5)

    ' Now find content between next > and </
    Dim openAngle, closeAngle
    openAngle = InStr(afterClose, ">")
    If openAngle = 0 Then
        ParseMetric = 0
        Exit Function
    End If
    closeAngle = InStr(Mid(afterClose, openAngle + 1), "<")
    If closeAngle = 0 Then
        ParseMetric = 0
        Exit Function
    End If

    rawVal = Mid(afterClose, openAngle + 1, closeAngle - 1)

    ' Strip spaces, currency symbols, non-numeric except . -
    cleanVal = CleanNumber(rawVal)
    If cleanVal = "" Then
        ParseMetric = 0
    Else
        On Error Resume Next
        ParseMetric = CDbl(cleanVal)
        If Err.Number <> 0 Then ParseMetric = 0
        On Error GoTo 0
    End If
End Function

' ============================================================
' ParsePLPeriod: find P/L for a named period row
' MT5 reports typically have rows labelled:
'   "Profit for today", "Profit for week", "Profit for month", "Profit for year"
' ============================================================
Function ParsePLPeriod(html, period)
    Dim lowerHtml, searchStr, pos, snippet
    lowerHtml = LCase(html)

    ' Try common MT5 label variants
    Dim labels(3)
    labels(0) = "profit for " & period
    labels(1) = "p/l for " & period
    labels(2) = "net profit " & period
    labels(3) = period & " profit"

    Dim j
    For j = 0 To 3
        pos = InStr(lowerHtml, labels(j))
        If pos > 0 Then
            ' Grab next 300 chars and parse number from next <td>
            snippet = Mid(html, pos, 300)
            Dim tdPos, valStart, valEnd, rawVal, cleanVal
            tdPos = InStr(LCase(snippet), "</td>")
            If tdPos > 0 Then
                Dim after
                after = Mid(snippet, tdPos + 5)
                Dim op, cp
                op = InStr(after, ">")
                If op > 0 Then
                    cp = InStr(Mid(after, op + 1), "<")
                    If cp > 0 Then
                        rawVal = Mid(after, op + 1, cp - 1)
                        cleanVal = CleanNumber(rawVal)
                        If cleanVal <> "" Then
                            On Error Resume Next
                            ParsePLPeriod = CDbl(cleanVal)
                            If Err.Number <> 0 Then ParsePLPeriod = 0
                            On Error GoTo 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next

    ParsePLPeriod = 0
End Function

' ============================================================
' CleanNumber: strip everything except digits, dot, minus sign
' Also handles numbers like "12 345.67" (space as thousands sep)
' ============================================================
Function CleanNumber(s)
    Dim result, c, i2
    result = ""
    s = Trim(s)
    For i2 = 1 To Len(s)
        c = Mid(s, i2, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        ElseIf c = "." Or c = "," Then
            ' Use dot as decimal separator, skip comma (thousands sep)
            If c = "." Then result = result & "."
        ElseIf c = "-" And i2 = 1 Then
            result = result & "-"
        End If
    Next
    CleanNumber = result
End Function

' ============================================================
' FormatNum: format number with 2 decimal places
' ============================================================
Function FormatNum(n)
    FormatNum = FormatNumber(n, 2, -1, 0, -1)
End Function

' ============================================================
' ColorClass: return CSS class based on value sign
' ============================================================
Function ColorClass(n)
    If n > 0 Then
        ColorClass = "profit"
    ElseIf n < 0 Then
        ColorClass = "loss"
    Else
        ColorClass = "neutral"
    End If
End Function

' ============================================================
' BuildDashboard: assemble full HTML string
' ============================================================
Function BuildDashboard(ids, bal, eq, plt, plw, plm, ply, found, _
                         tBal, tEq, tPLt, tPLw, tPLm, tPLy, ts)
    Dim h, j, rowClass, statusBadge

    h = "<!DOCTYPE html>" & vbCrLf
    h = h & "<html lang=""en"">" & vbCrLf
    h = h & "<head>" & vbCrLf
    h = h & "  <meta charset=""UTF-8"">" & vbCrLf
    h = h & "  <meta http-equiv=""refresh"" content=""300"">" & vbCrLf
    h = h & "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf
    h = h & "  <title>MT5 Multi-Account P/L Dashboard</title>" & vbCrLf
    h = h & "  <style>" & vbCrLf
    h = h & "    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }" & vbCrLf
    h = h & "    body { font-family: 'Segoe UI', Arial, sans-serif; background: #0d1117; color: #e6edf3; min-height: 100vh; padding: 20px; }" & vbCrLf
    h = h & "    h1 { font-size: 1.6rem; font-weight: 600; color: #58a6ff; margin-bottom: 4px; }" & vbCrLf
    h = h & "    .subtitle { font-size: 0.85rem; color: #8b949e; margin-bottom: 24px; }" & vbCrLf
    h = h & "    /* Summary cards */" & vbCrLf
    h = h & "    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 28px; }" & vbCrLf
    h = h & "    .card { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 16px 18px; }" & vbCrLf
    h = h & "    .card .label { font-size: 0.75rem; color: #8b949e; text-transform: uppercase; letter-spacing: .05em; margin-bottom: 6px; }" & vbCrLf
    h = h & "    .card .value { font-size: 1.35rem; font-weight: 700; }" & vbCrLf
    h = h & "    /* P/L color classes */" & vbCrLf
    h = h & "    .profit { color: #3fb950; }" & vbCrLf
    h = h & "    .loss   { color: #f85149; }" & vbCrLf
    h = h & "    .neutral { color: #e6edf3; }" & vbCrLf
    h = h & "    /* Table */" & vbCrLf
    h = h & "    .table-wrap { overflow-x: auto; }" & vbCrLf
    h = h & "    table { width: 100%; border-collapse: collapse; font-size: 0.9rem; }" & vbCrLf
    h = h & "    thead tr { background: #21262d; }" & vbCrLf
    h = h & "    th { padding: 10px 14px; text-align: right; color: #8b949e; font-weight: 600; font-size: 0.78rem; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }" & vbCrLf
    h = h & "    th:first-child { text-align: left; }" & vbCrLf
    h = h & "    td { padding: 10px 14px; text-align: right; border-bottom: 1px solid #21262d; white-space: nowrap; }" & vbCrLf
    h = h & "    td:first-child { text-align: left; font-weight: 600; color: #58a6ff; }" & vbCrLf
    h = h & "    tbody tr:hover { background: #1c2128; }" & vbCrLf
    h = h & "    tfoot tr { background: #21262d; font-weight: 700; }" & vbCrLf
    h = h & "    tfoot td { border-top: 2px solid #388bfd; border-bottom: none; color: #e6edf3; }" & vbCrLf
    h = h & "    tfoot td:first-child { color: #58a6ff; }" & vbCrLf
    h = h & "    .badge-ok      { display:inline-block; padding:1px 7px; border-radius:10px; font-size:0.7rem; background:#1a3a2a; color:#3fb950; border:1px solid #2ea043; }" & vbCrLf
    h = h & "    .badge-missing { display:inline-block; padding:1px 7px; border-radius:10px; font-size:0.7rem; background:#3a1a1a; color:#f85149; border:1px solid #b91c1c; }" & vbCrLf
    h = h & "    .footer { margin-top: 28px; font-size: 0.78rem; color: #8b949e; text-align: center; }" & vbCrLf
    h = h & "    section h2 { font-size: 1rem; color: #c9d1d9; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid #30363d; }" & vbCrLf
    h = h & "  </style>" & vbCrLf
    h = h & "</head>" & vbCrLf
    h = h & "<body>" & vbCrLf
    h = h & "  <h1>MT5 Multi-Account P/L Dashboard</h1>" & vbCrLf
    h = h & "  <p class=""subtitle"">Last updated: " & ts & " &nbsp;&bull;&nbsp; Auto-refresh every 5 minutes</p>" & vbCrLf

    ' ---- Summary cards ----
    h = h & "  <div class=""cards"">" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">Total Balance</div><div class=""value neutral"">" & FormatNum(tBal) & "</div></div>" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">Total Equity</div><div class=""value neutral"">" & FormatNum(tEq) & "</div></div>" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">P/L Today</div><div class=""value " & ColorClass(tPLt) & """>" & FormatNum(tPLt) & "</div></div>" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">P/L This Week</div><div class=""value " & ColorClass(tPLw) & """>" & FormatNum(tPLw) & "</div></div>" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">P/L This Month</div><div class=""value " & ColorClass(tPLm) & """>" & FormatNum(tPLm) & "</div></div>" & vbCrLf
    h = h & "    <div class=""card""><div class=""label"">P/L This Year</div><div class=""value " & ColorClass(tPLy) & """>" & FormatNum(tPLy) & "</div></div>" & vbCrLf
    h = h & "  </div>" & vbCrLf

    ' ---- Per-account table ----
    h = h & "  <section>" & vbCrLf
    h = h & "    <h2>Per-Account Breakdown</h2>" & vbCrLf
    h = h & "    <div class=""table-wrap"">" & vbCrLf
    h = h & "    <table>" & vbCrLf
    h = h & "      <thead><tr>" & vbCrLf
    h = h & "        <th>Account</th>" & vbCrLf
    h = h & "        <th>Status</th>" & vbCrLf
    h = h & "        <th>Balance</th>" & vbCrLf
    h = h & "        <th>Equity</th>" & vbCrLf
    h = h & "        <th>P/L Today</th>" & vbCrLf
    h = h & "        <th>P/L Week</th>" & vbCrLf
    h = h & "        <th>P/L Month</th>" & vbCrLf
    h = h & "        <th>P/L Year</th>" & vbCrLf
    h = h & "      </tr></thead>" & vbCrLf
    h = h & "      <tbody>" & vbCrLf

    For j = 0 To ACCOUNT_COUNT - 1
        If found(j) Then
            statusBadge = "<span class=""badge-ok"">OK</span>"
        Else
            statusBadge = "<span class=""badge-missing"">Missing</span>"
        End If

        h = h & "        <tr>" & vbCrLf
        h = h & "          <td>" & ids(j) & "</td>" & vbCrLf
        h = h & "          <td style=""text-align:center"">" & statusBadge & "</td>" & vbCrLf
        h = h & "          <td>" & FormatNum(bal(j)) & "</td>" & vbCrLf
        h = h & "          <td>" & FormatNum(eq(j)) & "</td>" & vbCrLf
        h = h & "          <td class=""" & ColorClass(plt(j)) & """>" & FormatNum(plt(j)) & "</td>" & vbCrLf
        h = h & "          <td class=""" & ColorClass(plw(j)) & """>" & FormatNum(plw(j)) & "</td>" & vbCrLf
        h = h & "          <td class=""" & ColorClass(plm(j)) & """>" & FormatNum(plm(j)) & "</td>" & vbCrLf
        h = h & "          <td class=""" & ColorClass(ply(j)) & """>" & FormatNum(ply(j)) & "</td>" & vbCrLf
        h = h & "        </tr>" & vbCrLf
    Next

    h = h & "      </tbody>" & vbCrLf
    h = h & "      <tfoot><tr>" & vbCrLf
    h = h & "        <td colspan=""2"">TOTAL</td>" & vbCrLf
    h = h & "        <td>" & FormatNum(tBal) & "</td>" & vbCrLf
    h = h & "        <td>" & FormatNum(tEq) & "</td>" & vbCrLf
    h = h & "        <td class=""" & ColorClass(tPLt) & """>" & FormatNum(tPLt) & "</td>" & vbCrLf
    h = h & "        <td class=""" & ColorClass(tPLw) & """>" & FormatNum(tPLw) & "</td>" & vbCrLf
    h = h & "        <td class=""" & ColorClass(tPLm) & """>" & FormatNum(tPLm) & "</td>" & vbCrLf
    h = h & "        <td class=""" & ColorClass(tPLy) & """>" & FormatNum(tPLy) & "</td>" & vbCrLf
    h = h & "      </tr></tfoot>" & vbCrLf
    h = h & "    </table>" & vbCrLf
    h = h & "    </div>" & vbCrLf
    h = h & "  </section>" & vbCrLf

    h = h & "  <p class=""footer"">MT5 Multi-Account P/L Dashboard &bull; Generated by merge_summary.vbs &bull; " & ts & "</p>" & vbCrLf
    h = h & "</body>" & vbCrLf
    h = h & "</html>" & vbCrLf

    BuildDashboard = h
End Function
