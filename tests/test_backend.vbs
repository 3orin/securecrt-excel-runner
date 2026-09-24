' Tests the CalcBook backend in rdwr.vbs against tests\test_book.xlsx (synthetic).
' Loads rdwr.vbs with ExecuteGlobal (Main is defined, not run - no SecureCRT needed).
' The test book is never saved; a round-trip copy goes to tests\out\.
' A LibreOffice window opens briefly. Usage: cscript //nologo tests\test_backend.vbs
Option Explicit
Dim fso, strDir, strBook, strOut, nPass, nFail
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
strBook = fso.BuildPath(strDir, "test_book.xlsx")
strOut = fso.BuildPath(strDir, "out")
If Not fso.FolderExists(strOut) Then fso.CreateFolder strOut
strOut = fso.BuildPath(strOut, "roundtrip.xlsx")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(strDir), "rdwr.vbs")).ReadAll()

Sub Check(strName, bOk, strGot)
	If bOk Then
		nPass = nPass + 1 : WScript.Echo "PASS  " & strName
	Else
		nFail = nFail + 1 : WScript.Echo "FAIL  " & strName & "   got [" & strGot & "]"
	End If
End Sub

Function Url(p)
	Url = "file:///" & Replace(Replace(p, "\", "/"), " ", "%20")
End Function

Dim sm, desk, b, b2, doc, sh, e, n0, n1, d, cell, ann, pv, filt, rt, rsh
nPass = 0 : nFail = 0
Set sm = CreateObject("com.sun.star.ServiceManager")
Set desk = sm.createInstance("com.sun.star.frame.Desktop")

Function CountDocs()
	Dim en, k : k = 0
	Set en = desk.getComponents().createEnumeration()
	Do While en.hasMoreElements() : en.nextElement() : k = k + 1 : Loop
	CountDocs = k
End Function

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

n0 = CountDocs()
Set b = NewSheetBackend("calc")
b.OpenBook strBook, 2
Set doc = FindDoc(Url(strBook))
Check "OpenBook opened the file", Not doc Is Nothing, ""
Set sh = doc.getSheets().getByIndex(1)
Check "sheet 2 is PromptResp", sh.getName() = "PromptResp", sh.getName()
Check "sheet 2 is active", doc.getCurrentController().getActiveSheet().getName() = "PromptResp", ""

' --- GetCell
Check "GetCell D2 text", b.GetCell(2, 4) = "echo alpha", b.GetCell(2, 4)
Check "GetCell B6 terminator", b.GetCell(6, 2) = "<< END >>", b.GetCell(6, 2)
Check "GetCell B3 skip marker", b.GetCell(3, 2) = "x", b.GetCell(3, 2)
Check "GetCell E2 number", b.GetCell(2, 5) = "42", b.GetCell(2, 5)
Check "GetCell empty F5", b.GetCell(5, 6) = "", b.GetCell(5, 6)
Check "GetCell far empty T20", b.GetCell(20, 20) = "", b.GetCell(20, 20)

' --- SetCell
b.SetCell 3, 8, "=SKiPPED="
Set cell = sh.getCellByPosition(7, 2)
Check "SetCell =SKiPPED= reads back", b.GetCell(3, 8) = "=SKiPPED=", b.GetCell(3, 8)
Check "SetCell =SKiPPED= is text, not formula", cell.getType() = 2, cell.getType()
b.SetCell 2, 8, "Ok"
Check "SetCell Ok", b.GetCell(2, 8) = "Ok", b.GetCell(2, 8)

' --- SetDate
d = Now
b.SetDate 2, 9, d
Set cell = sh.getCellByPosition(8, 1)
Check "SetDate stores a number", cell.getType() = 1, cell.getType()
Check "SetDate value = VBScript date", Abs(cell.getValue() - CDbl(d)) < 0.00001, cell.getValue()
Check "SetDate displays as date/time", InStr(cell.getString(), ":") > 0, cell.getString()

' --- SetNote
b.SetNote 2, 8, "line1" & vbCrLf & "line2", 300, 100
Set ann = sh.getCellByPosition(7, 1).getAnnotation()
Check "SetNote multi-line text", ann.getString() = "line1" & vbCrLf & "line2", ann.getString()
Check "SetNote width 300pt", Abs(ann.getAnnotationShape().getSize().Width - 10583) <= 2, ann.getAnnotationShape().getSize().Width
Check "SetNote height 100pt", Abs(ann.getAnnotationShape().getSize().Height - 3528) <= 2, ann.getAnnotationShape().getSize().Height
b.SetNote 2, 8, "replaced", 200, 100
Check "SetNote replaces, no duplicate", sh.getAnnotations().getCount() = 1, sh.getAnnotations().getCount()
Check "SetNote replaced text", sh.getCellByPosition(7, 1).getAnnotation().getString() = "replaced", ""
Check "SetNote width 200pt", Abs(sh.getCellByPosition(7, 1).getAnnotation().getAnnotationShape().getSize().Width - 7056) <= 2, ""
b.SetNote 3, 8, "stale from last run", 300, 100
Check "SetNote second cell", sh.getAnnotations().getCount() = 2, sh.getAnnotations().getCount()
b.SetNote 3, 8, "", 300, 100
Check "SetNote empty text clears stale note", sh.getAnnotations().getCount() = 1, sh.getAnnotations().getCount()
Check "SetNote empty text: other note untouched", sh.getCellByPosition(7, 1).getAnnotation().getString() = "replaced", ""

' --- attach to the already-open document instead of opening a second copy
n1 = CountDocs()
Set b2 = NewSheetBackend("calc")
b2.OpenBook strBook, 2
Check "second OpenBook attaches (no new doc)", CountDocs() = n1, CountDocs()
Check "second backend sees first backend's write", b2.GetCell(2, 8) = "Ok", b2.GetCell(2, 8)

' --- CloseBook releases only
b.CloseBook : b2.CloseBook
Check "CloseBook leaves document open", Not FindDoc(Url(strBook)) Is Nothing, ""
Check "CloseBook did not save", doc.isModified(), doc.isModified()

' --- round trip through .xlsx (how the evidence survives)
Set filt = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
filt.Name = "FilterName" : filt.Value = "Calc MS Excel 2007 XML"
doc.storeToURL Url(strOut), Array(filt)
Set pv = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
pv.Name = "Hidden" : pv.Value = True
Set rt = desk.loadComponentFromURL(Url(strOut), "_blank", 0, Array(pv))
Set rsh = rt.getSheets().getByIndex(1)
Check "roundtrip: H2 text", rsh.getCellByPosition(7, 1).getString() = "Ok", rsh.getCellByPosition(7, 1).getString()
' xlsx-imported notes load lazily: getString() is "" until the shape is touched, so read via the shape
Check "roundtrip: H2 note", rsh.getCellByPosition(7, 1).getAnnotation().getAnnotationShape().getString() = "replaced", rsh.getCellByPosition(7, 1).getAnnotation().getAnnotationShape().getString()
Check "roundtrip: H3 stale note gone", rsh.getAnnotations().getCount() = 1, rsh.getAnnotations().getCount()
Check "roundtrip: H3 still text", rsh.getCellByPosition(7, 2).getType() = 2, rsh.getCellByPosition(7, 2).getType()
Check "roundtrip: I2 date value", Abs(rsh.getCellByPosition(8, 1).getValue() - CDbl(d)) < 0.00002, rsh.getCellByPosition(8, 1).getValue()

' --- re-run over notes that came from an xlsx (the real workbook's situation)
Dim b3
Set b3 = NewSheetBackend("calc")
b3.OpenBook strOut, 2
b3.SetNote 2, 8, "rerun", 300, 100
Check "rerun: imported note replaced, no duplicate", rsh.getAnnotations().getCount() = 1, rsh.getAnnotations().getCount()
Check "rerun: new note text", rsh.getCellByPosition(7, 1).getAnnotation().getAnnotationShape().getString() = "rerun", rsh.getCellByPosition(7, 1).getAnnotation().getAnnotationShape().getString()
b3.CloseBook
rt.setModified False
rt.close True

' --- auto picks Calc when Excel is not running
Check "auto backend without Excel = CalcBook", TypeName(NewSheetBackend("auto")) = "CalcBook", TypeName(NewSheetBackend("auto"))

' --- cleanup: discard changes to the test book
doc.setModified False
doc.close True
' (not a count check: opening a file can replace LibreOffice's Start Center, which is also a component)
Check "test book closed", FindDoc(Url(strBook)) Is Nothing, CountDocs()

WScript.Echo vbCrLf & nPass & " passed, " & nFail & " failed"
WScript.Quit nFail
