' Builds tests\test_book.xlsx from scratch with LibreOffice Calc - synthetic data only.
' Same layout as the real driver workbook: sheet 1 HostList, sheet 2 PromptResp.
' Rows target SecureCRT Local Shell (cmd.exe) so step 3 can run against it.
' Usage: cscript //nologo tests\make_test_book.vbs
Option Explicit
Dim fso, strOut, sm, desk, pv, doc, sh, r, rows, i, j, filt

Set fso = CreateObject("Scripting.FileSystemObject")
strOut = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "test_book.xlsx")

Set sm = CreateObject("com.sun.star.ServiceManager")
Set desk = sm.createInstance("com.sun.star.frame.Desktop")
Set pv = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
pv.Name = "Hidden" : pv.Value = True
Set doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array(pv))

doc.getSheets().getByIndex(0).setName "HostList"
doc.getSheets().insertNewByName "PromptResp", 1
Set sh = doc.getSheets().getByIndex(1)

' A, B (skip / << END >>), C prompt, D command, E, F expected, G, H result, I date
rows = Array( _
	Array("", "",          "C:\>",  "Command", "", "Expected", "", "Result", "Date"), _
	Array("", "",          ">",     "echo alpha", "", "alpha", "", "", ""), _
	Array("", "x",         ">",     "echo must-not-run", "", "", "", "", ""), _
	Array("", "",          ">",     "ver", "", "Windows", "", "", ""), _
	Array("", "",          ">",     "echo no-expected-response", "", "", "", "", ""), _
	Array("", "<< END >>", "",      "", "", "", "", "", ""), _
	Array("", "",          "",      "notes below END must never be sent", "", "", "", "", ""), _
	Array("", "",          "",      "echo DANGER-past-end", "", "", "", "", "") _
)
For i = 0 To UBound(rows)
	For j = 0 To UBound(rows(i))
		If rows(i)(j) <> "" Then sh.getCellByPosition(j, i).setString rows(i)(j)
	Next
Next
sh.getCellByPosition(4, 1).setValue 42     ' E2: a numeric cell, to check GetCell on numbers

Set filt = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
filt.Name = "FilterName" : filt.Value = "Calc MS Excel 2007 XML"
doc.storeToURL "file:///" & Replace(Replace(strOut, "\", "/"), " ", "%20"), Array(filt)
doc.close True
WScript.Echo "wrote " & strOut
