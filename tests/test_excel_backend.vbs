' Tests the ExcelBook backend in rdwr.vbs - the MS Excel counterpart of test_backend.vbs.
' Builds a synthetic book with Excel COM (tests\out\excel_book.xlsx), then drives it through
' ExcelBook exactly as Main does. Loads rdwr.vbs with ExecuteGlobal (Main is defined, not run).
' Needs MS Excel and NO Excel running (the test starts its own, visible, and quits it at the end).
' If it aborts, close the Excel window it left without saving.
' Usage: cscript //nologo tests\test_excel_backend.vbs
Option Explicit
Dim fso, strDir, strOutDir, strBook, strRt, nPass, nFail
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
strOutDir = fso.BuildPath(strDir, "out")
If Not fso.FolderExists(strOutDir) Then fso.CreateFolder strOutDir
strBook = fso.BuildPath(strOutDir, "excel_book.xlsx")
strRt = fso.BuildPath(strOutDir, "excel_roundtrip.xlsx")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(strDir), "rdwr.vbs")).ReadAll()

Sub Check(strName, bOk, strGot)
	If bOk Then
		nPass = nPass + 1 : WScript.Echo "PASS  " & strName
	Else
		nFail = nFail + 1 : WScript.Echo "FAIL  " & strName & "   got [" & strGot & "]"
	End If
End Sub

' CR / LF made visible, for messages
Function Show(s)
	Show = Replace(Replace(s, vbCr, "\r"), vbLf, "\n")
End Function

Function BigText(nBytes)
	Dim s, k : s = "" : k = 0
	Do While Len(s) < nBytes
		k = k + 1
		s = s & "ge-0/0/" & Right("00" & k, 3) & "  up    up   inet  10.0." & (k Mod 250) & ".1/24" & vbCrLf
	Loop
	BigText = Left(s, nBytes)
End Function

Dim xl, e, errText
nPass = 0 : nFail = 0

' --- Excel must not be running: the test must not touch the user's own workbooks
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
e = Err.Number
On Error Goto 0
If e = 0 Then
	WScript.Echo "Excel is already running - close it (save your work) and run the test again."
	WScript.Quit 99
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = True
WScript.Echo "Excel " & xl.Version & " build " & xl.Build & vbCrLf

' --- build the synthetic book (same layout as make_test_book.vbs, plus Settings)
Dim wb, sh, rows, i, j
If fso.FileExists(strBook) Then fso.DeleteFile strBook
If fso.FileExists(strRt) Then fso.DeleteFile strRt
Set wb = xl.Workbooks.Add
Do While wb.Sheets.Count < 3 : wb.Sheets.Add , wb.Sheets(wb.Sheets.Count) : Loop
Do While wb.Sheets.Count > 3 : xl.DisplayAlerts = False : wb.Sheets(4).Delete : xl.DisplayAlerts = True : Loop
wb.Sheets(1).Name = "HostList"
wb.Sheets(2).Name = "PromptResp"
wb.Sheets(3).Name = "Settings"
Set sh = wb.Sheets(2)
rows = Array( _
	Array("", "",          "C:\>",  "Command", "", "Expected", "", "Result", "Date"), _
	Array("", "",          ">",     "echo alpha", "", "alpha", "", "", ""), _
	Array("", "x",         ">",     "echo must-not-run", "", "", "", "", ""), _
	Array("", "",          ">",     "ver", "", "Windows", "", "", ""), _
	Array("", "",          ">",     "echo no-expected-response", "", "", "", "", ""), _
	Array("", "<< END >>", "",      "", "", "", "", "", ""), _
	Array("", "",          "",      "notes below END must never be sent", "", "", "", "", "") _
)
For i = 0 To UBound(rows)
	For j = 0 To UBound(rows(i))
		If rows(i)(j) <> "" Then sh.Cells(i + 1, j + 1).Value = rows(i)(j)
	Next
Next
sh.Cells(2, 5).Value = 42          ' E2: numeric cell
sh.Cells(3, 5).Formula = "=NA()"   ' E3: formula error cell (#N/A)
wb.Sheets(1).Activate              ' OpenBook must switch to sheet 2 itself
With wb.Sheets(3)
	.Cells(1, 1).Value = "Key" : .Cells(1, 2).Value = "Value"
	.Cells(2, 1).Value = "TIMEOUT"   : .Cells(2, 2).Value = 5
	.Cells(3, 1).Value = "PAGER"     : .Cells(3, 2).Value = "'---(more"
	.Cells(4, 1).Value = "PAGER_KEY" : .Cells(4, 2).Value = "SPACE"
	.Cells(5, 1).Value = "ERROR"     : .Cells(5, 2).Value = "is ambiguous"
End With
xl.DisplayAlerts = False
wb.SaveAs strBook, 51               ' xlOpenXMLWorkbook
xl.DisplayAlerts = True
wb.Close False
Set wb = Nothing : Set sh = Nothing
Check "test book built", fso.FileExists(strBook), strBook

' --- OpenBook: book not open yet -> opens it, selects sheet 2
Dim b, b2, n0, cell, d, s, c
n0 = xl.Workbooks.Count
Set b = NewSheetBackend("excel")
b.OpenBook strBook, 2
Check "OpenBook opened the file", xl.Workbooks.Count = n0 + 1, xl.Workbooks.Count
Set wb = xl.Workbooks("excel_book.xlsx")
Set sh = wb.Sheets(2)
Check "sheet 2 is PromptResp", sh.Name = "PromptResp", sh.Name
Check "sheet 2 is active", xl.ActiveSheet.Name = "PromptResp", xl.ActiveSheet.Name

' --- GetCell
Check "GetCell D2 text", b.GetCell(2, 4) = "echo alpha", b.GetCell(2, 4)
Check "GetCell B6 terminator", b.GetCell(6, 2) = "<< END >>", b.GetCell(6, 2)
Check "GetCell B3 skip marker", b.GetCell(3, 2) = "x", b.GetCell(3, 2)
Check "GetCell E2 number", b.GetCell(2, 5) = "42", b.GetCell(2, 5)
Check "GetCell empty F5", b.GetCell(5, 6) = "", b.GetCell(5, 6)
Check "GetCell far empty T20", b.GetCell(20, 20) = "", b.GetCell(20, 20)
On Error Resume Next
s = b.GetCell(3, 5)
e = Err.Number : errText = Err.Description
On Error Goto 0
Check "GetCell on a #N/A cell does not raise", e = 0, "error " & e & " " & errText
Check "GetCell on a #N/A cell = #N/A (as Calc)", s = "#N/A", s

' --- GetNamedSheetCell + LoadSettings (Settings sheet found by name)
Check "GetNamedSheetCell Settings A2", b.GetNamedSheetCell("Settings", 2, 1) = "TIMEOUT", b.GetNamedSheetCell("Settings", 2, 1)
Check "GetNamedSheetCell number B2", b.GetNamedSheetCell("Settings", 2, 2) = "5", b.GetNamedSheetCell("Settings", 2, 2)
Check "GetNamedSheetCell apostrophe-typed B3", b.GetNamedSheetCell("Settings", 3, 2) = "---(more", b.GetNamedSheetCell("Settings", 3, 2)
Check "GetNamedSheetCell name is case-insensitive", b.GetNamedSheetCell("SETTINGS", 2, 1) = "TIMEOUT", b.GetNamedSheetCell("SETTINGS", 2, 1)
Check "GetNamedSheetCell missing sheet = empty", b.GetNamedSheetCell("NoSuchSheet", 2, 1) = "", b.GetNamedSheetCell("NoSuchSheet", 2, 1)
Set g_book = b
LoadSettings
Check "LoadSettings TIMEOUT 5", g_TIMEOUT_SEC = 5, g_TIMEOUT_SEC
Check "LoadSettings PAGER replaces defaults", UBound(g_vPager) = 0 And g_vPager(0) = "---(more", Join(g_vPager, "|")
Check "LoadSettings PAGER_KEY SPACE", g_strPagerKey = " ", "[" & g_strPagerKey & "]"
Check "LoadSettings ERROR replaces defaults", UBound(g_vErrors) = 0 And g_vErrors(0) = "is ambiguous", Join(g_vErrors, "|")

' --- SetCell
b.SetCell 3, 8, "=SKiPPED="
Set cell = sh.Cells(3, 8)
Check "SetCell =SKiPPED= reads back", b.GetCell(3, 8) = "=SKiPPED=", b.GetCell(3, 8)
Check "SetCell =SKiPPED= is text, not formula", Not cell.HasFormula And VarType(cell.Value) = vbString, "HasFormula=" & cell.HasFormula & " Formula=" & cell.Formula
Check "SetCell =SKiPPED= shows as typed", cell.Text = "=SKiPPED=", cell.Text
b.SetCell 2, 8, "Ok"
Check "SetCell Ok", b.GetCell(2, 8) = "Ok", b.GetCell(2, 8)
b.SetCell 4, 8, "FAiL: error text '% Invalid'"
Check "SetCell FAiL with quotes", b.GetCell(4, 8) = "FAiL: error text '% Invalid'", b.GetCell(4, 8)
b.SetCell 5, 8, "-5 dBm"
Check "SetCell leading '-' stays text", b.GetCell(5, 8) = "-5 dBm" And Not sh.Cells(5, 8).HasFormula, b.GetCell(5, 8)

' --- SetDate
d = Now
b.SetDate 2, 9, d
Set cell = sh.Cells(2, 9)
Check "SetDate stores a date", VarType(cell.Value) = vbDate, "VarType " & VarType(cell.Value)
Check "SetDate value = VBScript date", Abs(cell.Value2 - CDbl(d)) < 0.00001, cell.Value2
Check "SetDate formatted as date + time", InStr(cell.NumberFormat, "h") > 0 And InStr(cell.NumberFormat, "y") > 0, cell.NumberFormat
' Excel does not widen the column for a COM write (it does when typed): a narrow I shows ####
WScript.Echo "      (I2 displays [" & cell.Text & "] at column width " & cell.ColumnWidth & ")"

' --- SetNote
b.SetNote 2, 8, "line1" & vbCrLf & "line2", 300, 100
Set c = sh.Cells(2, 8).Comment
Check "SetNote creates a comment", Not c Is Nothing, ""
Check "SetNote multi-line text", c.Text = "line1" & vbCrLf & "line2", Show(c.Text)
Check "SetNote width 300pt", Abs(c.Shape.Width - 300) <= 1, c.Shape.Width
Check "SetNote height 100pt", Abs(c.Shape.Height - 100) <= 1, c.Shape.Height
WScript.Echo "      (comment author = [" & c.Author & "] - Application.UserName, saved in the file)"
b.SetNote 2, 8, "replaced", 200, 100
Check "SetNote replaces, no duplicate", sh.Comments.Count = 1, sh.Comments.Count
Check "SetNote replaced text", sh.Cells(2, 8).Comment.Text = "replaced", Show(sh.Cells(2, 8).Comment.Text)
Check "SetNote width 200pt", Abs(sh.Cells(2, 8).Comment.Shape.Width - 200) <= 1, sh.Cells(2, 8).Comment.Shape.Width
b.SetNote 3, 8, "stale from last run", 300, 100
Check "SetNote second cell", sh.Comments.Count = 2, sh.Comments.Count
On Error Resume Next
b.SetNote 3, 8, "", 300, 100
e = Err.Number : errText = Err.Description
On Error Goto 0
Check "SetNote empty text does not raise", e = 0, "error " & e & " " & errText
s = "(none)"
If Not sh.Cells(3, 8).Comment Is Nothing Then s = sh.Cells(3, 8).Comment.Text
Check "SetNote empty text clears stale note (no 'user:' note)", s = "(none)", Show(s)
Check "SetNote empty text: one note left", sh.Comments.Count = 1, sh.Comments.Count
Check "SetNote empty text: other note untouched", sh.Cells(2, 8).Comment.Text = "replaced", ""

' long notes: terse output is ~2.6 KB; Excel's limit for a comment is 32,767 characters
Dim strBig
strBig = BigText(2600)
b.SetNote 4, 8, strBig, 450, 300
Check "SetNote 2.6 KB exact", sh.Cells(4, 8).Comment.Text = strBig, Len(sh.Cells(4, 8).Comment.Text) & " chars"
strBig = BigText(30000)
On Error Resume Next
b.SetNote 5, 8, strBig, 450, 300
e = Err.Number : errText = Err.Description
On Error Goto 0
Check "SetNote 30 KB does not raise", e = 0, "error " & e & " " & errText
If e = 0 Then Check "SetNote 30 KB exact", sh.Cells(5, 8).Comment.Text = strBig, Len(sh.Cells(5, 8).Comment.Text) & " chars"
strBig = BigText(40000)
On Error Resume Next
b.SetNote 6, 8, strBig, 450, 300
e = Err.Number : errText = Err.Description
On Error Goto 0
Check "SetNote 40 KB (over Excel's 32,767) does not raise", e = 0, "error " & e & " " & errText
If e = 0 Then
	s = sh.Cells(6, 8).Comment.Text
	Check "SetNote 40 KB fits the limit", Len(s) <= 32767, Len(s) & " chars"
	Check "SetNote 40 KB keeps the start", Left(s, 32687) = Left(strBig, 32687), ""
	errText = vbCrLf & "[... 7313 chars cut - Excel note limit 32767]"
	Check "SetNote 40 KB ends with the cut marker", Right(s, Len(errText)) = errText, Show(Right(s, 60))
End If

' --- attach: a second OpenBook on the open (modified) book must not reopen or prompt
Dim n1
n1 = xl.Workbooks.Count
Set b2 = NewSheetBackend("excel")
b2.OpenBook strBook, 2
Check "second OpenBook attaches (no new book)", xl.Workbooks.Count = n1, xl.Workbooks.Count
Check "second backend sees first backend's write", b2.GetCell(2, 8) = "Ok", b2.GetCell(2, 8)

' --- auto picks Excel when it is running
Check "auto backend with Excel running = ExcelBook", TypeName(NewSheetBackend("auto")) = "ExcelBook", TypeName(NewSheetBackend("auto"))

' --- CloseBook releases only
b.CloseBook : b2.CloseBook
Set g_book = Nothing
Check "CloseBook leaves the book open", xl.Workbooks.Count = n1, xl.Workbooks.Count
Check "CloseBook did not save", Not wb.Saved, "Saved=" & wb.Saved

' --- round trip through .xlsx (how the evidence survives)
Dim rt, rsh, b3
wb.SaveCopyAs strRt
Set rt = xl.Workbooks.Open(strRt)
Set rsh = rt.Sheets(2)
Check "roundtrip: H2 text", rsh.Cells(2, 8).Text = "Ok", rsh.Cells(2, 8).Text
Check "roundtrip: H2 note", rsh.Cells(2, 8).Comment.Text = "replaced", Show(rsh.Cells(2, 8).Comment.Text)
Check "roundtrip: H3 still text", Not rsh.Cells(3, 8).HasFormula And rsh.Cells(3, 8).Value = "=SKiPPED=", rsh.Cells(3, 8).Formula
Check "roundtrip: H4 2.6 KB note", rsh.Cells(4, 8).Comment.Text = BigText(2600), Len(rsh.Cells(4, 8).Comment.Text) & " chars"
Check "roundtrip: I2 date value", Abs(rsh.Cells(2, 9).Value2 - CDbl(d)) < 0.00002, rsh.Cells(2, 9).Value2

' --- re-run over comments that came from an xlsx (the real workbook's situation)
Set b3 = NewSheetBackend("excel")
b3.OpenBook strRt, 2
b3.SetNote 2, 8, "rerun", 300, 100
Check "rerun: imported note replaced", rsh.Cells(2, 8).Comment.Text = "rerun", Show(rsh.Cells(2, 8).Comment.Text)
b3.CloseBook

' --- cleanup: nothing saved; quit the Excel this test started
rt.Close False
wb.Close False
Set c = Nothing : Set cell = Nothing : Set sh = Nothing : Set rsh = Nothing
Set wb = Nothing : Set rt = Nothing
xl.Quit
Set xl = Nothing

WScript.Echo vbCrLf & nPass & " passed, " & nFail & " failed"
WScript.Quit nFail
