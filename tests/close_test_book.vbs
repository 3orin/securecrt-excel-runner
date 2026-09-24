' Closes tests\test_book.xlsx in LibreOffice without saving (cleanup after an aborted test run).
' Touches no other document.
Dim fso, u, cu, sm, desk, en, c, n
Set fso = CreateObject("Scripting.FileSystemObject")
u = LCase("file:///" & Replace(Replace(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "test_book.xlsx"), "\", "/"), " ", "%20"))
Set sm = CreateObject("com.sun.star.ServiceManager")
Set desk = sm.createInstance("com.sun.star.frame.Desktop")
Set en = desk.getComponents().createEnumeration()
n = 0
Do While en.hasMoreElements()
	Set c = en.nextElement()
	cu = ""
	On Error Resume Next
	cu = LCase(c.getURL())
	On Error Goto 0
	If cu = u Then c.setModified False : c.close True : n = n + 1
Loop
WScript.Echo "closed " & n & " copy(ies) of test_book.xlsx"
