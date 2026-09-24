' ExcelSpreadsheets-ReadingAndWriting.vbs
' following general format:
'          A       |   B   |     C    |     D    |     E     |    F    |  G   |    H    |      I              |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+---------------------+
' 1 |              |       | r-cor-01#| sh run   |           | hostname|      | Succ    | 22/02/22  16:35:00  |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+---------------------+
' 2 |              |       | r-cor-01#| ping a.b |           | !!!!!   |      |  Succ   |                     |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+---------------------+
' 3 |              |       | r-cor-01#| conf t   |           | (conf)# |      |  Fail   |                     |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+---------------------+
' 4 |              |       | (conf)#  | exit     |           |  r-01#  |      |  Succ   |                     |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+---------------------+
'
Dim g_shell
Set g_shell = CreateObject("WSCript.Shell")

' File path [full]
Dim g_strSpreadSheetPath
g_strSpreadSheetPath = "C:\securecrt-excel-runner\samples\sample_cisco_devnet.xlsx"

' Spreadsheet backend: "excel" (MS Office), "calc" (LibreOffice) or "auto"
' auto = use Excel if it is running, otherwise LibreOffice Calc
Dim g_BACKEND
g_BACKEND = "auto"

Dim g_book

Dim g_SPARE_COL, g_INACT_COL, g_CMD_COL, g_RESULT1_COL, g_ACT_COL

' Convert Letter column indicators to numerical references
g_SPARE_COL        = Asc("A") - 64
g_ACT_COL          = Asc("B") - 64
g_PROMPT_COL       = Asc("C") - 64
g_CMD_COL          = Asc("D") - 64
g_RESULT1_COL      = Asc("F") - 64
g_RESULT2_COL      = Asc("G") - 64
g_ACT_RES_COL      = Asc("H") - 64
g_ACT_DATE_COL     = Asc("I") - 64
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 2 DO / 2 ADD:
' Check XL formula substitution [ + INDIRECT ()]
' Check Linux CLi: prompt + responces, PiNG + NC [per DEST + PER PORT, from every HOST]
' 0prgMany version
' manout = spit out the string  with N ms delay
' MANY output patameters = like OLD ProcommPlus script [until empty cell]
' ADO DB connection: https://www.softwaretestinghelp.com/vbscript-connection-objects-tutorial-12/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Main()
	' workbook with PROMPT, COMMANDS & RESPONSES must be OPEN already (Excel)
	Set g_book = NewSheetBackend(g_BACKEND)
	' Script sheet = always 2nd
	g_book.OpenBook g_strSpreadSheetPath, 2
	LoadSettings

    Dim nRowIndex
    ' Skip the header [1st] row:
    nRowIndex = 2

    ' Loop continues until an "END" row, or a row with B, C and D all empty
    Dim strPrompt, strCmd, strResp1, strActive, strStatus, strOutput, nTimeout, nTimedOutRow
    nTimedOutRow = 0
    Do
        strActive = Trim(g_book.GetCell(nRowIndex, g_ACT_COL))
		' "<< END >>" in any spacing / case ends the run
		If UCase(Replace(Replace(strActive, " ", ""), vbTab, "")) = "<<END>>" Then Exit Do
		strPrompt = Trim(g_book.GetCell(nRowIndex, g_PROMPT_COL))
		strCmd   =  Trim(g_book.GetCell(nRowIndex, g_CMD_COL))
		' an empty row ends the run too - it is never sent to the device
		If strActive = "" And strPrompt = "" And strCmd = "" Then Exit Do
        If strActive = "" Then
			' ECLu HE KOMMEHT TO
			' expected result + optional per-row timeout (column G, seconds)
			strResp1 =  Trim(g_book.GetCell(nRowIndex, g_RESULT1_COL))
			nTimeout = Trim(g_book.GetCell(nRowIndex, g_RESULT2_COL))
			If IsNumeric(nTimeout) And nTimeout <> "" Then nTimeout = CLng(nTimeout) Else nTimeout = g_TIMEOUT_SEC
			If nTimeout < 1 Then nTimeout = g_TIMEOUT_SEC

			' The last row that ran timed out: its command may still be running, or waiting at a
			' question ([confirm]). Anything typed now would be swallowed or taken as the answer,
			' so first get this row's prompt back. No prompt within this row's timeout: stop.
			If nTimedOutRow > 0 Then
				If Not WaitForPrompt(strPrompt, nTimeout) Then
					g_book.SetCell nRowIndex, g_ACT_RES_COL, "FAiL: not sent - no prompt " & nTimeout & "s after row " & nTimedOutRow & " timed out"
					g_book.SetNote nRowIndex, g_ACT_RES_COL, "[not sent - SecureCRT screen at that moment:]" & vbCrLf & ScreenText(""), 200, 100
					g_book.SetDate nRowIndex, g_ACT_DATE_COL, Now
					Exit Do
				End If
			End If

                ' Heuristically determine the shell's prompt
                Do
                    ' Simulate pressing "Enter" so the prompt appears again...
					' KZ: ETOT CRLF mym = HyzhEH, 6E3 HEGO XPEHOBO PA6OTAET
                      crt.Screen.Send vbcr
                    ' Attempt to detect the command prompt heuristically by
                    ' waiting for the cursor to stop moving... (the timeout for
                    ' WaitForCursor above might not be enough for slower-
                    ' responding hosts, so you will need to adjust the timeout
                    ' value above to meet your system's specific timing
                    ' requirements).
                    Do
                        bCursorMoved = crt.Screen.WaitForCursor(1)
                    Loop Until bCursorMoved = False
                    ' Once the cursor has stopped moving for about a second,
                    ' we'll assume it's safe to start interacting with the
                    ' remote system. Get the shell prompt so that we can know
                    ' what to look for when determining if the command is
                    ' completed. Won't work if the prompt is dynamic (e.g.,
                    ' changes according to current working folder, etc.)
                    nRow = crt.Screen.CurrentRow
                    strPrintout = crt.screen.Get(nRow, _
                                               0, _
                                               nRow, _
                                               crt.Screen.CurrentColumn - 1)
                    ' Loop until we actually see a line of text appear:
                    strPrintout = Trim(strPrintout)
                    If strPrintout <> "" Then Exit Do
                Loop

			strStatus = RunStep(strPrompt, strCmd, strResp1, nTimeout, strOutput)
			If Left(strStatus, 13) = "FAiL: timeout" Then nTimedOutRow = nRowIndex Else nTimedOutRow = 0
			g_book.SetCell nRowIndex, g_ACT_RES_COL, strStatus
			If strStatus = "Ok" Then
				g_book.SetNote nRowIndex, g_ACT_RES_COL, strOutput, 300, 100
			Else
				g_book.SetNote nRowIndex, g_ACT_RES_COL, strOutput, 200, 100
			End If
 		Else
          ' mark the skipped ones in the spreadsheet
           g_book.SetCell nRowIndex, g_ACT_RES_COL, "=SKiPPED="

       End If
       ' We always record the date of action status
       g_book.SetDate nRowIndex, g_ACT_DATE_COL, Now

        ' move down to the next row in the spreadsheet
        nRowIndex = nRowIndex + 1
    Loop

 '== Workbook is never saved or closed - to prevent accidental XL file locking
 ' User will have to save & close file himself
    g_book.CloseBook
    Set g_book = Nothing

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' One row: send the command, then read until the prompt comes back.
'   F tokens: 1st = success, any other = failure. Reading carries on past the
'   match to the prompt, so the note holds the whole output.
'   F empty: Ok when the prompt returns, unless an error marker is in the output.
'   Pager prompts are answered with the pager key and left out of the note.
'   Every read has a timeout, so a missing prompt can not hang the run.
' Returns the status for column H; strOutput gets the text for the note.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function RunStep(strPrompt, strCmd, strResp1, nTimeout, ByRef strOutput)
	Dim vTokens, vWait, nTok, nPromptIdx, i, strChunk, nIdx, strStatus, bAfterPager, strWord, bTimedOut

	' expected tokens, empty ones dropped ("a  b" must not wait for "")
	vTokens = Array()
	For Each strWord In Split(strResp1, " ")
		If strWord <> "" Then
			ReDim Preserve vTokens(UBound(vTokens) + 1)
			vTokens(UBound(vTokens)) = strWord
		End If
	Next
	nTok = UBound(vTokens) + 1

	' wait list = tokens, then the prompt, then the pager strings (MatchIndex is 1-based)
	vWait = vTokens
	nPromptIdx = 0
	If strPrompt <> "" Then
		ReDim Preserve vWait(UBound(vWait) + 1)
		vWait(UBound(vWait)) = strPrompt
		nPromptIdx = UBound(vWait) + 1
	End If
	For i = 0 To UBound(g_vPager)
		ReDim Preserve vWait(UBound(vWait) + 1)
		vWait(UBound(vWait)) = g_vPager(i)
	Next

	' Synchronous while this step reads: otherwise data that arrives between two reads
	' (e.g. the rest of the line after an F token) is shown on screen but never reaches
	' the script. Switched on only now - the prompt detection above has just waited for
	' the screen to settle, so nothing old is queued - and off again at the end, so
	' leftovers are not carried into the next row.
	crt.Screen.Synchronous = True

	' Issue command(s) to remote machine - OTOCLAL OTBET , row = CURRENT
	crt.Screen.Send strCmd & vbcr

	strOutput = ""
	strStatus = ""
	bAfterPager = False
	bTimedOut = False
	Do
		strChunk = crt.Screen.ReadString(vWait, nTimeout)
		' MatchIndex = which string was found; 0 = timeout
		nIdx = crt.Screen.MatchIndex
		If bAfterPager Then strChunk = StripPagerTail(strChunk)
		bAfterPager = False
		If nIdx > nTok And nIdx <> nPromptIdx Then strChunk = StripPagerLead(strChunk)
		strOutput = strOutput & strChunk

		If nIdx = 0 Then
			If strStatus = "" Then
				strStatus = "FAiL: timeout " & nTimeout & "s"
				bTimedOut = True
			Else
				strOutput = strOutput & vbCrLf & "[no prompt within " & nTimeout & "s]"
			End If
			Exit Do
		ElseIf nIdx > nTok And nIdx <> nPromptIdx Then
			' pager: advance a page, drop the pager prompt from the note
			crt.Screen.Send g_strPagerKey
			bAfterPager = True
		ElseIf nIdx = nPromptIdx Then
			If strStatus = "" Then
				If nTok = 0 Then
					strStatus = ErrorMarkerStatus(strOutput)
				Else
					strStatus = "FAiL: expected text not seen"
				End If
			End If
			Exit Do
		Else
			' an F token: keep the matched text in the note
			strOutput = strOutput & vTokens(nIdx - 1)
			If strStatus = "" Then
				If nIdx = 1 Then strStatus = "Ok" Else strStatus = "FAiL: got " & vTokens(nIdx - 1)
			End If
			' no prompt to read up to: stop at the token
			If nPromptIdx = 0 Then Exit Do
		End If
	Loop

	crt.Screen.Synchronous = False

	' a timed-out ReadString returns nothing, so keep what the screen showed instead
	If bTimedOut Then
		strOutput = "[timeout " & nTimeout & "s - SecureCRT screen at that moment:]" & vbCrLf & ScreenText(strCmd) & strOutput
	End If

	' the prompt itself is never in the capture; drop the line break before it
	Do While Len(strOutput) > 0 And InStr(vbCr & vbLf, Right(strOutput, 1)) > 0
		strOutput = Left(strOutput, Len(strOutput) - 1)
	Loop
	RunStep = strStatus
End Function

' After a timed-out row: press Enter and wait for strPrompt. An idle device answers at once;
' a busy one (IOS drops the Enter while a ping runs) once it has finished. Called before the
' 1-second settle in Main, which then soaks up any extra prompt a type-ahead device prints for
' the queued Enter. No prompt in column C: nothing to wait for.
Function WaitForPrompt(strPrompt, nTimeout)
	WaitForPrompt = True
	If strPrompt = "" Then Exit Function
	crt.Screen.Synchronous = True
	crt.Screen.Send vbCr
	crt.Screen.ReadString Array(strPrompt), nTimeout
	WaitForPrompt = (crt.Screen.MatchIndex <> 0)
	crt.Screen.Synchronous = False
End Function

' The visible SecureCRT screen from the last line that ends with strCmd (this row's
' command echo) down - older rows and earlier runs still on screen are left out.
' Command not on screen (scrolled off, wrapped, or empty): the whole screen.
' Trailing blanks cut per line, empty lines at the bottom dropped.
Function ScreenText(strCmd)
	Dim vLines, i, nFirst, nLast, s
	vLines = Split(Replace(crt.Screen.Get2(1, 1, crt.Screen.Rows, crt.Screen.Columns), vbCrLf, vbLf), vbLf)
	nFirst = 0
	nLast = -1
	For i = 0 To UBound(vLines)
		vLines(i) = RTrim(vLines(i))
		If vLines(i) <> "" Then nLast = i
		If strCmd <> "" And Len(vLines(i)) >= Len(strCmd) Then
			If Right(vLines(i), Len(strCmd)) = strCmd Then nFirst = i
		End If
	Next
	s = ""
	For i = nFirst To nLast
		s = s & vLines(i) & vbCrLf
	Next
	ScreenText = s
End Function

' "Ok", or a FAiL naming the first error marker found after the echoed command line
Function ErrorMarkerStatus(strOutput)
	Dim strBody, nEol, i
	strBody = strOutput
	nEol = InStr(strBody, vbLf)
	If nEol > 0 Then strBody = Mid(strBody, nEol + 1) Else strBody = ""
	ErrorMarkerStatus = "Ok"
	For i = 0 To UBound(g_vErrors)
		If g_vErrors(i) <> "" Then
			If InStr(1, strBody, g_vErrors(i), vbTextCompare) > 0 Then
				ErrorMarkerStatus = "FAiL: error text '" & g_vErrors(i) & "'"
				Exit Function
			End If
		End If
	Next
End Function

' After a pager prompt matched (by prefix), the next read starts with the rest of the
' prompt line and the CR / backspace / space sequence the device uses to erase it.
' Remove both; a CR that starts a CR LF is the next line's (a blank line), so it stays.
' No erase at all: the line break after the prompt goes instead. The indentation of the
' next real line is kept.
' Junos: "---(more 31%)---" CR, 40 spaces, CR.  IOS: " --More-- " 9 BS, 9 spaces, 9 BS.
Function StripPagerTail(strChunk)
	Dim re
	Set re = New RegExp
	re.Pattern = "^[^\r\n\x08]*(?:(?:\x08|\r(?!\n))+ *)*(?:\x08|\r(?!\n))+|^[^\r\n\x08]*\r?\n"
	StripPagerTail = re.Replace(strChunk, "")
End Function

' The read that ends at a pager prompt: blanks on the prompt's own line, before the
' matched text, belong to the prompt (IOS " --More-- ") - drop them.
Function StripPagerLead(strChunk)
	Dim re
	Set re = New RegExp
	re.Pattern = "(^|\n) +$"
	StripPagerLead = re.Replace(strChunk, "$1")
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Settings: defaults below, overridden by an optional sheet named "Settings":
'   column A = key, column B = value, from row 2 down to the first empty key.
'   TIMEOUT    seconds per read (a number in column G of a step overrides it)
'   PAGER      one row per pager prompt, matched as a prefix, e.g. ---(more
'   PAGER_KEY  SPACE (default), ENTER, or literal text
'   ERROR      one row per error marker, case-insensitive
' Any PAGER / ERROR rows replace that whole default list.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim g_TIMEOUT_SEC, g_vPager, g_strPagerKey, g_vErrors

Sub LoadSettings()
	Dim nRow, strKey, strVal, strPagers, strErrors
	g_TIMEOUT_SEC = 30
	g_vPager = Array("---(more", "--More--", "<--- More --->", "-- More --", "---- More ----")
	g_strPagerKey = " "
	g_vErrors = Array("syntax error", "unknown command", "error:", "% Invalid", "% Incomplete", "% Ambiguous", "% Unknown")

	strPagers = "" : strErrors = ""
	For nRow = 2 To 200
		strKey = UCase(Trim(g_book.GetNamedSheetCell("Settings", nRow, 1)))
		If strKey = "" Then Exit For
		strVal = g_book.GetNamedSheetCell("Settings", nRow, 2)
		Select Case strKey
			Case "TIMEOUT"
				If IsNumeric(strVal) Then
					If CLng(strVal) > 0 Then g_TIMEOUT_SEC = CLng(strVal)
				End If
			Case "PAGER"
				If Trim(strVal) <> "" Then strPagers = strPagers & vbLf & strVal
			Case "PAGER_KEY"
				Select Case UCase(Trim(strVal))
					Case "SPACE", "" : g_strPagerKey = " "
					Case "ENTER"     : g_strPagerKey = vbCr
					Case Else        : g_strPagerKey = strVal
				End Select
			Case "ERROR"
				If Trim(strVal) <> "" Then strErrors = strErrors & vbLf & Trim(strVal)
		End Select
	Next
	If strPagers <> "" Then g_vPager = Split(Mid(strPagers, 2), vbLf)
	If strErrors <> "" Then g_vErrors = Split(Mid(strErrors, 2), vbLf)
End Sub

' DONE: simplify Prompt + Response
' DONE: XL = activate ONLY, no save / close / quit
' DONE: add TiME into comments
' DONE: RESULT = array; split on many , only 1st = success
' DONE: RESULT comparison; 1st = OK , others = HA XEP

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Spreadsheet backends. Both classes expose the same methods, so Main does not
' care which one it has:
'   OpenBook path, sheetNo       attach to (or open) the workbook, select sheet (1-based)
'   GetCell(row, col)            cell text, "" when empty
'   GetNamedSheetCell(name, r, c) same, on another sheet by name; "" if no such sheet
'   SetCell row, col, text       always literal text, never parsed as a formula
'   SetDate row, col, date       date/time value
'   SetNote row, col, text, w, h replace the cell comment; w/h in points
'   CloseBook                    release only - never saves, closes or quits
' Rows and columns are 1-based, as in Excel.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NewSheetBackend(strKind)
	Select Case LCase(strKind)
		Case "excel" : Set NewSheetBackend = New ExcelBook
		Case "calc"  : Set NewSheetBackend = New CalcBook
		Case Else
			' auto: Excel only if it is already running
			Dim objTry
			On Error Resume Next
			Set objTry = GetObject(, "Excel.Application")
			If Err.Number = 0 Then
				Set NewSheetBackend = New ExcelBook
			Else
				Set NewSheetBackend = New CalcBook
			End If
			On Error Goto 0
			Set objTry = Nothing
	End Select
End Function

Class ExcelBook
	Private m_app, m_book, m_sheet

	Public Sub OpenBook(strPath, nSheet)
		Set m_app = GetObject(, "Excel.Application")
		' if MS Excel ^^ app is NOT OPENED - script gives error HERE
		Set m_book = m_app.Workbooks.Open(strPath)
		' g_objExcel.ActiveWindow.WindowState = xlNormal <== Does NOT work
		m_app.Application.Visible = True
		Set m_sheet = m_book.Sheets(nSheet)
		m_sheet.Activate
	End Sub

	Public Function GetCell(nRow, nCol)
		GetCell = CellText(m_sheet.Cells(nRow, nCol))
	End Function

	' an error value (#N/A, #REF!...) makes CStr fail - return what the cell shows, as Calc does
	Private Function CellText(objCell)
		Dim v
		v = objCell.Value
		If VarType(v) = vbError Then
			CellText = CStr(objCell.Text)
		Else
			CellText = CStr(v)
		End If
	End Function

	Public Function GetNamedSheetCell(strSheet, nRow, nCol)
		Dim objSh
		GetNamedSheetCell = ""
		For Each objSh In m_book.Sheets
			If LCase(objSh.Name) = LCase(strSheet) Then
				GetNamedSheetCell = CellText(objSh.Cells(nRow, nCol))
				Exit Function
			End If
		Next
	End Function

	Public Sub SetCell(nRow, nCol, strText)
		Dim s
		s = CStr(strText)
		' leading apostrophe = Excel text prefix, so "=SKiPPED=" is not a formula
		If Len(s) > 0 Then
			If InStr("=+-@'", Left(s, 1)) > 0 Then s = "'" & s
		End If
		m_sheet.Cells(nRow, nCol).Value = s
	End Sub

	Public Sub SetDate(nRow, nCol, dtValue)
		m_sheet.Cells(nRow, nCol).Value = dtValue
	End Sub

	Public Sub SetNote(nRow, nCol, strText, nWidth, nHeight)
		Dim objCell, s, nKeep
		Set objCell = m_sheet.Cells(nRow, nCol)
		objCell.ClearComments
		' empty text: no note (AddComment "" would write the "UserName:" prefix)
		If CStr(strText) = "" Then Exit Sub
		' Excel keeps at most 32767 chars of a note and drops the rest silently - mark the cut
		s = CStr(strText)
		If Len(s) > 32767 Then
			nKeep = 32767 - 80
			s = Left(s, nKeep) & vbCrLf & "[... " & (Len(s) - nKeep) & " chars cut - Excel note limit 32767]"
		End If
		objCell.AddComment s
		objCell.Comment.Shape.Width = nWidth
		objCell.Comment.Shape.Height = nHeight
	End Sub

	Public Sub CloseBook
		' objWkBook.Save / objWkBook.Close / g_objExcel.Quit - deliberately not done
		Set m_sheet = Nothing
		Set m_book = Nothing
		Set m_app = Nothing
	End Sub
End Class

Class CalcBook
	Private m_sm, m_doc, m_sheet

	Public Sub OpenBook(strPath, nSheet)
		Dim objDesk, objEnum, objComp, strUrl, strCompUrl, noArgs
		Set m_sm = CreateObject("com.sun.star.ServiceManager")
		Set objDesk = m_sm.createInstance("com.sun.star.frame.Desktop")
		strUrl = LCase(PathToUrl(strPath))
		' use the document if it is already open in LibreOffice
		Set objEnum = objDesk.getComponents().createEnumeration()
		Do While objEnum.hasMoreElements()
			Set objComp = objEnum.nextElement()
			' not every component has getURL (e.g. the Start Center). Read it on its own
			' line: under On Error Resume Next a failing If condition runs its Then branch
			strCompUrl = ""
			On Error Resume Next
			strCompUrl = LCase(objComp.getURL())
			On Error Goto 0
			If strCompUrl = strUrl Then
				Set m_doc = objComp
				Exit Do
			End If
		Loop
		If Not IsObject(m_doc) Then
			noArgs = Array()
			Set m_doc = objDesk.loadComponentFromURL(strUrl, "_default", 0, noArgs)
		End If
		m_doc.getCurrentController().getFrame().getContainerWindow().setVisible True
		Set m_sheet = m_doc.getSheets().getByIndex(nSheet - 1)
		m_doc.getCurrentController().setActiveSheet m_sheet
	End Sub

	Public Function GetCell(nRow, nCol)
		GetCell = m_sheet.getCellByPosition(nCol - 1, nRow - 1).getString()
	End Function

	Public Function GetNamedSheetCell(strSheet, nRow, nCol)
		Dim objSheets, i
		GetNamedSheetCell = ""
		Set objSheets = m_doc.getSheets()
		For i = 0 To objSheets.getCount() - 1
			If LCase(objSheets.getByIndex(i).getName()) = LCase(strSheet) Then
				GetNamedSheetCell = objSheets.getByIndex(i).getCellByPosition(nCol - 1, nRow - 1).getString()
				Exit Function
			End If
		Next
	End Function

	Public Sub SetCell(nRow, nCol, strText)
		' setString stores literal text; no apostrophe prefix needed
		m_sheet.getCellByPosition(nCol - 1, nRow - 1).setString CStr(strText)
	End Sub

	Public Sub SetDate(nRow, nCol, dtValue)
		Dim objCell, objLocale
		Set objCell = m_sheet.getCellByPosition(nCol - 1, nRow - 1)
		' VBScript Date and Calc share the 1899-12-30 serial epoch
		objCell.setValue CDbl(dtValue)
		Set objLocale = m_sm.Bridge_GetStruct("com.sun.star.lang.Locale")
		objCell.NumberFormat = m_doc.getNumberFormats().getStandardFormat(6, objLocale) ' 6 = DATETIME
	End Sub

	Public Sub SetNote(nRow, nCol, strText, nWidth, nHeight)
		Dim objCell, objSize
		Set objCell = m_sheet.getCellByPosition(nCol - 1, nRow - 1)
		' insertNew replaces an existing note on the same cell; "" removes it
		' (Excel would keep an empty comment - either way there is no evidence)
		m_sheet.getAnnotations().insertNew objCell.getCellAddress(), CStr(strText)
		If Len(strText) = 0 Then Exit Sub
		' points -> 1/100 mm
		Set objSize = m_sm.Bridge_GetStruct("com.sun.star.awt.Size")
		objSize.Width = CLng(nWidth * 2540 / 72)
		objSize.Height = CLng(nHeight * 2540 / 72)
		objCell.getAnnotation().getAnnotationShape().setSize objSize
	End Sub

	Public Sub CloseBook
		' document stays open and unsaved in LibreOffice, same as the Excel backend
		Set m_sheet = Nothing
		Set m_doc = Nothing
		Set m_sm = Nothing
	End Sub

	Private Function PathToUrl(strPath)
		Dim s
		s = Replace(strPath, "\", "/")
		s = Replace(s, "%", "%25")
		s = Replace(s, " ", "%20")
		s = Replace(s, "#", "%23")
		PathToUrl = "file:///" & s
	End Function
End Class
