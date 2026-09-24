' Runs rdwr.vbs Main against tests\fake_crt.vbs and a synthetic workbook, then checks
' every status and note. No SecureCRT, no device.
' Backend: argument "calc" (LibreOffice, a window opens briefly) or "excel" (MS Excel - must NOT be
' running; the test starts its own and quits it). No argument: calc if LibreOffice is installed, else excel.
' Usage: cscript //nologo tests\test_logic.vbs [calc|excel]
Option Explicit
Dim fso, strDir, crt, nPass, nFail, sm, desk, xl, strBackend, e
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(strDir, "fake_crt.vbs")).ReadAll()
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(strDir), "rdwr.vbs")).ReadAll()
strBackend = ""
If WScript.Arguments.Count > 0 Then strBackend = LCase(WScript.Arguments(0))
If strBackend <> "excel" Then
	On Error Resume Next
	Set sm = CreateObject("com.sun.star.ServiceManager")
	e = Err.Number
	On Error Goto 0
	If e = 0 Then
		strBackend = "calc"
		Set desk = sm.createInstance("com.sun.star.frame.Desktop")
	ElseIf strBackend = "calc" Then
		WScript.Echo "LibreOffice not found." : WScript.Quit 99
	Else
		strBackend = "excel"
	End If
End If
If strBackend = "excel" Then
	On Error Resume Next
	Set xl = GetObject(, "Excel.Application")
	e = Err.Number
	On Error Goto 0
	' Main attaches with GetObject - it must find the Excel this test starts, not the user's
	If e = 0 Then WScript.Echo "Excel is already running - close it (save your work) and run the test again." : WScript.Quit 99
	Set xl = CreateObject("Excel.Application")
	xl.Visible = True
	WScript.Echo "backend: excel (Excel " & xl.Version & " build " & xl.Build & ")"
Else
	WScript.Echo "backend: calc (LibreOffice)"
End If
nPass = 0 : nFail = 0

Const P = "admin@lab-srx>"

Sub Check(strName, bOk, strGot)
	If bOk Then
		nPass = nPass + 1 : WScript.Echo "PASS  " & strName
	Else
		nFail = nFail + 1 : WScript.Echo "FAIL  " & strName & vbCrLf & "      got [" & strGot & "]"
	End If
End Sub

' notes are compared on content: any line break -> LF
Function NL(s)
	NL = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

Function Url(p)
	Url = "file:///" & Replace(Replace(p, "\", "/"), " ", "%20")
End Function

Function Pv(n, v)
	Set Pv = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	Pv.Name = n : Pv.Value = v
End Function

' rows: Array(B, C, D, F, G); vSettings: Empty or Array of Array(key, value)
Sub MakeBook(strPath, vRows, vSettings)
	If strBackend = "excel" Then MakeBookExcel strPath, vRows, vSettings Else MakeBookCalc strPath, vRows, vSettings
End Sub

' every cell as text, like Calc's setString: the apostrophe keeps "---(more" / "5" from being parsed
Sub MakeBookExcel(strPath, vRows, vSettings)
	Dim wb, sh, i, j, st
	If fso.FileExists(strPath) Then fso.DeleteFile strPath
	Set wb = xl.Workbooks.Add
	Do While wb.Sheets.Count < 2 : wb.Sheets.Add , wb.Sheets(wb.Sheets.Count) : Loop
	xl.DisplayAlerts = False
	Do While wb.Sheets.Count > 2 : wb.Sheets(3).Delete : Loop
	xl.DisplayAlerts = True
	wb.Sheets(1).Name = "HostList"
	wb.Sheets(2).Name = "PromptResp"
	Set sh = wb.Sheets(2)
	sh.Cells(1, 2).Value = "B" : sh.Cells(1, 3).Value = "PROMPT"
	sh.Cells(1, 4).Value = "COMMAND" : sh.Cells(1, 6).Value = "RESPONSES"
	For i = 0 To UBound(vRows)
		For j = 0 To 4
			If vRows(i)(j) <> "" Then sh.Cells(i + 2, Array(2, 3, 4, 6, 7)(j)).Value = "'" & vRows(i)(j)
		Next
	Next
	If IsArray(vSettings) Then
		Set st = wb.Sheets.Add(, wb.Sheets(2))
		st.Name = "Settings"
		st.Cells(1, 1).Value = "KEY" : st.Cells(1, 2).Value = "VALUE"
		For i = 0 To UBound(vSettings)
			st.Cells(i + 2, 1).Value = "'" & vSettings(i)(0)
			st.Cells(i + 2, 2).Value = "'" & vSettings(i)(1)
		Next
	End If
	wb.Sheets(1).Activate
	xl.DisplayAlerts = False
	wb.SaveAs strPath, 51      ' xlOpenXMLWorkbook
	xl.DisplayAlerts = True
	wb.Close False
End Sub

Sub MakeBookCalc(strPath, vRows, vSettings)
	Dim doc, sh, i, j, st
	Set doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array(Pv("Hidden", True)))
	doc.getSheets().getByIndex(0).setName "HostList"
	doc.getSheets().insertNewByName "PromptResp", 1
	Set sh = doc.getSheets().getByIndex(1)
	sh.getCellByPosition(1, 0).setString "B" : sh.getCellByPosition(2, 0).setString "PROMPT"
	sh.getCellByPosition(3, 0).setString "COMMAND" : sh.getCellByPosition(5, 0).setString "RESPONSES"
	For i = 0 To UBound(vRows)
		For j = 0 To 4
			If vRows(i)(j) <> "" Then sh.getCellByPosition(Array(1, 2, 3, 5, 6)(j), i + 1).setString vRows(i)(j)
		Next
	Next
	If IsArray(vSettings) Then
		doc.getSheets().insertNewByName "Settings", 2
		Set st = doc.getSheets().getByIndex(2)
		st.getCellByPosition(0, 0).setString "KEY" : st.getCellByPosition(1, 0).setString "VALUE"
		For i = 0 To UBound(vSettings)
			st.getCellByPosition(0, i + 1).setString vSettings(i)(0)
			st.getCellByPosition(1, i + 1).setString vSettings(i)(1)
		Next
	End If
	doc.storeToURL Url(strPath), Array(Pv("FilterName", "Calc MS Excel 2007 XML"))
	doc.close True
End Sub

Function FindDoc(u)
	Dim en, c, cu
	Set FindDoc = Nothing
	Set en = desk.getComponents().createEnumeration()
	Do While en.hasMoreElements()
		Set c = en.nextElement()
		cu = ""
		On Error Resume Next
		cu = LCase(c.getURL())
		On Error Goto 0
		If cu = LCase(u) Then Set FindDoc = c
	Loop
End Function

' results of the run: H text and H note per row, from the (unsaved) open document
Dim g_res, g_note, g_date
Sub Collect(strPath, nRows)
	If strBackend = "excel" Then CollectExcel strPath, nRows Else CollectCalc strPath, nRows
End Sub

Sub CollectExcel(strPath, nRows)
	Dim wb, sh, r
	Set wb = xl.Workbooks(fso.GetFileName(strPath))
	Set sh = wb.Sheets(2)
	ReDim g_res(nRows + 1) : ReDim g_note(nRows + 1) : ReDim g_date(nRows + 1)
	For r = 2 To nRows + 1
		g_res(r) = CStr(sh.Cells(r, 8).Value)
		g_date(r) = CStr(sh.Cells(r, 9).Value)
		g_note(r) = ""
		If Not sh.Cells(r, 8).Comment Is Nothing Then g_note(r) = sh.Cells(r, 8).Comment.Text
		g_note(r) = NL(g_note(r))
	Next
	wb.Close False
End Sub

Sub CollectCalc(strPath, nRows)
	Dim doc, sh, r
	Set doc = FindDoc(Url(strPath))
	Set sh = doc.getSheets().getByIndex(1)
	ReDim g_res(nRows + 1) : ReDim g_note(nRows + 1) : ReDim g_date(nRows + 1)
	For r = 2 To nRows + 1
		g_res(r) = sh.getCellByPosition(7, r - 1).getString()
		g_date(r) = sh.getCellByPosition(8, r - 1).getString()
		g_note(r) = ""
		On Error Resume Next
		g_note(r) = sh.getCellByPosition(7, r - 1).getAnnotation().getAnnotationShape().getString()
		On Error Goto 0
		g_note(r) = NL(g_note(r))
	Next
	doc.setModified False
	doc.close True
End Sub

Function RunBook(strName, vRows, vSettings)
	Dim strPath
	strPath = fso.BuildPath(strDir, "out")
	If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
	strPath = fso.BuildPath(strPath, strName)
	MakeBook strPath, vRows, vSettings
	Set crt = New FakeCrt
	If g_ios Then crt.Screen.UseIOS
	SetOutputs crt.Screen
	g_strSpreadSheetPath = strPath
	g_BACKEND = strBackend
	Main
	Collect strPath, UBound(vRows) + 1
	RunBook = strPath
End Function

' synthetic device outputs
Dim g_terse, g_blanky, g_ios
g_ios = False
Sub SetOutputs(scr)
	Dim i, s
	' 4 lines per page at length 5: pages 2 and 3 start on a blank line, page 4 on an indented one
	g_blanky = Join(Array("L1", "L2", "L3", "L4", "", "L6", "L7", "L8", "", "", "L11", "L12", "  indented 13", "L14"), vbCrLf)
	scr.SetOutput "show blanky", g_blanky
	scr.SetOutput "show clock", "*11:05:51.358 UTC Thu Sep 24 2026"
	scr.SetOutput "show version", "Hostname: lab-srx" & vbCrLf & "Model: srx300" & vbCrLf & "Junos: 23.4R2-S3.9" & vbCrLf & "JUNOS Software Release [23.4R2-S3.9]"
	scr.SetOutput "show system uptime", "Current time: 2026-09-24 02:12:07 EST" & vbCrLf & "Time Source:  NTP CLOCK " & vbCrLf & _
		"System booted: 2026-09-22 11:20:30 EST (1d 14:51 ago)" & vbCrLf & " 2:12AM  up 1 day, 14:52, 2 users, load averages: 0.06, 0.05, 0.01"
	s = "Interface               Admin Link Proto    Local                 Remote"
	For i = 0 To 30
		s = s & vbCrLf & "ge-0/0/" & i & Space(24 - Len("ge-0/0/" & i)) & "up    up"
		If i Mod 7 = 3 Then s = s & vbCrLf & Space(35) & "inet6   "
	Next
	s = s & vbCrLf & "vtep                    up    up"
	g_terse = s
	scr.SetOutput "show interfaces terse", s
	scr.SetOutput "show log wibble", "Sep 24 02:00:00 something wibble happened"
End Sub

Dim strEcho
'=============================================================================
WScript.Echo "--- book A: book 1 + book 2 + timeouts, default settings"
RunBook "logic_a.xlsx", Array( _
	Array("", P, "set cli screen-length 0", "Screen", ""), _
	Array("", P, "show version", "Model:", ""), _
	Array("", P, "show system uptime", "", ""), _
	Array("", P, "show bogus-command", "", ""), _
	Array("", P, "show bogus-command", "NOSUCH syntax", ""), _
	Array("", P, "show system uptime", "NOSUCH", ""), _
	Array("x", P, "show system users", "", ""), _
	Array("", P, "set cli screen-length 10", "Screen", ""), _
	Array("", P, "show interfaces terse", "vtep", ""), _
	Array("", P, "set cli screen-length 0", "Screen", ""), _
	Array("", "WRONG-PROMPT>", "show system uptime", "NOSUCH", "5"), _
	Array("", P, "show system uptime", "booted:", ""), _
	Array("", P, "hang", "", ""), _
	Array("", P, "show interfaces terse", "", ""), _
	Array("<< end >>", "", "", "", ""), _
	Array("", P, "show version | match DANGER-PAST-END", "", "") _
), Empty

Check "A2  Ok", g_res(2) = "Ok", g_res(2)
Check "A2  #5 fixed: reply is in the note", g_note(2) = "set cli screen-length 0" & vbLf & "Screen length set to 0", g_note(2)
Check "A3  Ok on 1st token", g_res(3) = "Ok", g_res(3)
Check "A3  #5 fixed: note complete past the token", InStr(g_note(3), "Hostname: lab-srx" & vbLf & "Model: srx300" & vbLf & "Junos: 23.4R2") > 0 And Right(g_note(3), 13) = "[23.4R2-S3.9]" , g_note(3)
Check "A4  empty F, prompt back = Ok", g_res(4) = "Ok", g_res(4)
Check "A4  #5 fixed: last line intact", Right(g_note(4), 31) = "load averages: 0.06, 0.05, 0.01", g_note(4)
Check "A5  #3 fixed: error text fails", g_res(5) = "FAiL: error text 'syntax error'", g_res(5)
Check "A5  note keeps the caret line", InStr(g_note(5), "                    ^") > 0, g_note(5)
Check "A6  2nd token = FAiL", g_res(6) = "FAiL: got syntax", g_res(6)
Check "A7  expected text missing = FAiL", g_res(7) = "FAiL: expected text not seen", g_res(7)
Check "A8  skip", g_res(8) = "=SKiPPED=", g_res(8)
Check "A9  pager on", g_res(9) = "Ok", g_res(9)
Check "A10 #4 fixed: paged command Ok", g_res(10) = "Ok", g_res(10)
Check "A10 #4 fixed: no pager text in note", InStr(g_note(10), "(more") = 0, g_note(10)
Check "A10 #4 fixed: paged note = unpaged output", g_note(10) = NL("show interfaces terse" & vbCrLf & g_terse), g_note(10)
Check "A15 unpaged reference = same text", g_note(15) = g_note(10), g_note(15)
Check "A11 pager off", g_res(11) = "Ok", g_res(11)
Check "A12 #1 fixed: wrong prompt times out (G = 5)", g_res(12) = "FAiL: timeout 5s", g_res(12)
Check "A13 #1 fixed: run carried on", g_res(13) = "Ok", g_res(13)
Check "A13 rest of the matched line kept (synchronous read)", InStr(g_note(13), "System booted: 2026-09-22 11:20:30 EST (1d 14:51 ago)" & vbLf) > 0, g_note(13)
Check "A   synchronous mode switched off again after each row", crt.Screen.Synchronous = False, crt.Screen.Synchronous
Check "A14 #1 fixed: no prompt at all = default timeout", g_res(14) = "FAiL: timeout 30s", g_res(14)
Check "A12 timeout note = screen snapshot", InStr(g_note(12), "[timeout 5s - SecureCRT screen at that moment:]" & vbLf) = 1, g_note(12)
Check "A12 snapshot shows what the device printed", InStr(g_note(12), "System booted: 2026-09-22") > 0 And InStr(g_note(12), "admin@lab-srx>") > 0, g_note(12)
Check "A12 snapshot lines have no trailing blanks", InStr(g_note(12), " " & vbLf) = 0, g_note(12)
Check "A12 snapshot starts at this row's command", InStr(g_note(12), "moment:]" & vbLf & "admin@lab-srx> show system uptime" & vbLf) > 0, g_note(12)
Check "A12 snapshot leaves out older rows", InStr(g_note(12), "Screen length set to") = 0 And InStr(g_note(12), "ge-0/0/") = 0, g_note(12)
Check "A14 hung command: snapshot shows it", InStr(g_note(14), "[timeout 30s") = 1 And InStr(g_note(14), "moment:]" & vbLf & "admin@lab-srx> hang" & vbLf & "working...") > 0, g_note(14)
Check "A14 snapshot leaves out older rows", InStr(g_note(14), "System booted") = 0, g_note(14)
Check "A16 #2 fixed: '<< end >>' ends the run", g_res(16) = "" And g_date(16) = "", g_res(16)
Check "A17 never sent", InStr(crt.Screen.SentLog, "DANGER") = 0, "sent"
Check "A   pager answered with Space", InStr(crt.Screen.SentLog, "terse" & vbCr & " ") > 0, ""
Check "A   dates written", g_date(2) <> "" And g_date(15) <> "", g_date(2)

'=============================================================================
WScript.Echo "--- book B: terminator typo"
RunBook "logic_b.xlsx", Array( _
	Array("", P, "show version", "Junos:", ""), _
	Array("<<END>>", "", "", "", ""), _
	Array("", P, "show version | match DANGER-PAST-END", "", "") _
), Empty
Check "B2  Ok", g_res(2) = "Ok", g_res(2)
Check "B3  #2 fixed: '<<END>>' ends the run", g_res(3) = "", g_res(3)
Check "B4  never sent", InStr(crt.Screen.SentLog, "DANGER") = 0, "sent"

'=============================================================================
WScript.Echo "--- book C: empty row ends the run"
RunBook "logic_c.xlsx", Array( _
	Array("", P, "show version", "Junos:", ""), _
	Array("#", "", "", "", ""), _
	Array("", "", "", "", ""), _
	Array("", P, "show version | match DANGER-PAST-END", "", "") _
), Empty
Check "C2  Ok", g_res(2) = "Ok", g_res(2)
Check "C3  '#' row with nothing else is skipped, not the end", g_res(3) = "=SKiPPED=", g_res(3)
Check "C4  #2 fixed: empty row not run", g_res(4) = "" And g_date(4) = "", g_res(4)
Check "C5  never sent", InStr(crt.Screen.SentLog, "DANGER") = 0, "sent"

'=============================================================================
WScript.Echo "--- book D: Settings sheet"
RunBook "logic_d.xlsx", Array( _
	Array("", P, "show log wibble", "", ""), _
	Array("", P, "show bogus-command", "", ""), _
	Array("", P, "hang", "", ""), _
	Array("", P, "hang", "", "12"), _
	Array("", P, "set cli screen-length 10", "Screen", ""), _
	Array("", P, "show interfaces terse", "", ""), _
	Array("", P, "hangbig", "", "") _
), Array(Array("TIMEOUT", "7"), Array("ERROR", "wibble"), Array("PAGER", "---(more"), Array("PAGER_KEY", "ENTER"))
Check "D2  custom error marker fails", g_res(2) = "FAiL: error text 'wibble'", g_res(2)
Check "D3  ERROR rows replace the defaults", g_res(3) = "Ok", g_res(3)
Check "D4  TIMEOUT from Settings", g_res(4) = "FAiL: timeout 7s", g_res(4)
Check "D5  column G overrides it", g_res(5) = "FAiL: timeout 12s", g_res(5)
Check "D7  PAGER_KEY ENTER still pages to the end", g_res(7) = "Ok" And InStr(g_note(7), "(more") = 0, g_res(7)
Check "D7  ... and the note is complete", g_note(7) = NL("show interfaces terse" & vbCrLf & g_terse), g_note(7)
Check "D8  command scrolled off: whole screen kept", InStr(g_note(8), "[timeout 7s") = 1 And InStr(g_note(8), "line 30") > 0 And InStr(g_note(8), "hangbig") = 0, g_note(8)
Check "D8  ... which is at most one screen (24 rows)", UBound(Split(g_note(8), vbLf)) <= 24, UBound(Split(g_note(8), vbLf))

'=============================================================================
WScript.Echo "--- book E: Cisco IOS pager (live 2026-09-24: leading blank, lost blank lines)"
Const PI = "Cat8kv#"
g_ios = True
RunBook "logic_e.xlsx", Array( _
	Array("", PI, "show blanky", "", ""), _
	Array("", PI, "terminal length 5", "", ""), _
	Array("", PI, "show blanky", "", ""), _
	Array("", PI, "show blanky", "L14", "") _
), Empty
g_ios = False
Check "E2  unpaged reference", g_res(2) = "Ok" And g_note(2) = NL("show blanky" & vbCrLf & g_blanky), g_note(2)
Check "E4  paged --More-- Ok", g_res(4) = "Ok", g_res(4)
Check "E4  no pager text or backspaces in the note", InStr(g_note(4), "More") = 0 And InStr(g_note(4), Chr(8)) = 0, g_note(4)
Check "E4  paged note = unpaged (blank before --More-- dropped, blank lines kept)", g_note(4) = g_note(2), g_note(4)
Check "E5  F token on the last page", g_res(5) = "Ok" And g_note(5) = g_note(2), g_note(5)

'=============================================================================
WScript.Echo "--- book F: Junos pager, pages starting on blank lines"
RunBook "logic_f.xlsx", Array( _
	Array("", P, "set cli screen-length 5", "Screen", ""), _
	Array("", P, "show blanky", "", "") _
), Empty
Check "F3  paged note = unpaged, blank lines kept", g_res(3) = "Ok" And g_note(3) = NL("show blanky" & vbCrLf & g_blanky), g_note(3)

'=============================================================================
WScript.Echo "--- book G: row timed out while the command still runs (IOS drops type-ahead)"
g_ios = True
RunBook "logic_g.xlsx", Array( _
	Array("", PI, "busy 10", "!!!! .....", "3"), _
	Array("", PI, "show clock", "UTC", ""), _
	Array("", PI, "show clock", "UTC", "") _
), Empty
g_ios = False
Check "G2  still running at the timeout", g_res(2) = "FAiL: timeout 3s", g_res(2)
Check "G3  next row waited for the prompt, then ran", g_res(3) = "Ok", g_res(3)
Check "G3  its note is its own output only", g_note(3) = NL("show clock" & vbCrLf & "*11:05:51.358 UTC Thu Sep 24 2026"), g_note(3)
Check "G4  run in step afterwards", g_res(4) = "Ok" And g_note(4) = g_note(3), g_note(4)

'=============================================================================
WScript.Echo "--- book H: the same on a device that queues type-ahead (stray prompts)"
RunBook "logic_h.xlsx", Array( _
	Array("", P, "busyq 10", "!!!! .....", "3"), _
	Array("", P, "show system uptime", "booted:", ""), _
	Array("", P, "show version", "Junos:", "") _
), Empty
Check "H2  still running at the timeout", g_res(2) = "FAiL: timeout 3s", g_res(2)
Check "H3  waited, then ran - not judged on the ping's prompt", g_res(3) = "Ok", g_res(3)
Check "H3  note holds no ping leftovers", InStr(g_note(3), "Success rate") = 0 And InStr(g_note(3), "System booted") > 0, g_note(3)
Check "H4  no stray prompt shifted the next row", g_res(4) = "Ok" And InStr(g_note(4), "Model: srx300") > 0, g_note(4)

'=============================================================================
WScript.Echo "--- book I: the prompt never comes back after a timeout - stop, send nothing"
g_ios = True
RunBook "logic_i.xlsx", Array( _
	Array("", PI, "dead", "!!!! .....", "3"), _
	Array("x", PI, "show users", "", ""), _
	Array("", PI, "show clock", "UTC", "4"), _
	Array("", PI, "show version | include DANGER-AFTER-STOP", "", "") _
), Empty
g_ios = False
Check "I2  timeout", g_res(2) = "FAiL: timeout 3s", g_res(2)
Check "I3  skip rows are still just skipped", g_res(3) = "=SKiPPED=", g_res(3)
Check "I4  not sent: no prompt within its own timeout", g_res(4) = "FAiL: not sent - no prompt 4s after row 2 timed out", g_res(4)
Check "I4  note = screen snapshot", InStr(g_note(4), "[not sent - SecureCRT screen at that moment:]") = 1 And InStr(g_note(4), "ICMP Echos") > 0, g_note(4)
Check "I4  its command was never typed", InStr(crt.Screen.SentLog, "show clock") = 0, crt.Screen.SentLog
Check "I5  run stopped there", g_res(5) = "" And g_date(5) = "" And InStr(crt.Screen.SentLog, "DANGER") = 0, g_res(5)

If strBackend = "excel" Then xl.Quit : Set xl = Nothing
WScript.Echo vbCrLf & nPass & " passed, " & nFail & " failed  (backend: " & strBackend & ")"
WScript.Quit nFail
