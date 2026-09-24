' Fake SecureCRT "crt" object for offline tests of rdwr.vbs. Synthetic device only.
' Behaves like the Junos SRX seen live on 2026-09-24:
'   - echoes the command, prints canned output, then the prompt
'   - "set cli screen-length N" turns the pager on (N-1 lines per page)
'   - pager prompt "---(more)---" then "---(more NN%)---"; Space = next page,
'     erased with CR + 40 spaces + CR
'   - unknown command -> caret line + "syntax error, expecting <command>."
'   - "hang" never returns the prompt; "hangbig" prints 30 lines first (echo scrolls off)
' UseIOS switches to Cisco IOS-XE as seen live 2026-09-24 (DevNet Cat8kv): prompt "Cat8kv#",
'   "terminal length N" (no reply), pager " --More-- " erased with 9 BS + 9 spaces + 9 BS.
' Like the real thing in non-synchronous mode, WaitForCursor throws away unread data,
' and so does ReadString: after a match, the rest of that line (it arrived in the same
' packet, before the next read) is lost - seen live 2026-09-24. Synchronous = True keeps it.
' ReadString(list, timeout): earliest match wins; "" matches at once (as seen live);
' no match = timeout: MatchIndex 0 and an EMPTY string - what was read is discarded
' (seen live 2026-09-24, book 3 row 26).
' Time: a fake clock. A read that times out moves it on by its timeout, WaitForCursor(n) by n.
' Busy commands - a no-reply ping, as seen live on IOS 2026-09-24 (rows 28-29):
'   "busy N"  prints the ping header and one ".", finishes N seconds later (rest + prompt);
'             whatever is typed meanwhile is DROPPED, like IOS during a ping
'   "busyq N" the same, but typed input is QUEUED and run afterwards (type-ahead devices)
'   "dead"    busy for ever, input dropped
' A read that could match once the busy command finishes waits for it, if within its timeout.

Class FakeScreen
	Public MatchIndex, CurrentRow, CurrentColumn, Synchronous, Rows, Columns
	Public SentLog, LastTimeout, ReadCount
	Private m_pending, m_shown, m_prompt, m_screenLen, m_left, m_total, m_out, m_ios
	Private m_now, m_busyUntil, m_busyRest, m_busyQueue, m_queued

	Private Sub Class_Initialize()
		m_prompt = "admin@lab-srx> "
		m_pending = "" : m_shown = "" : m_screenLen = 0 : m_left = Empty : Rows = 24 : Columns = 80
		SentLog = "" : MatchIndex = 0 : ReadCount = 0 : Synchronous = False : m_ios = False
		m_now = 0 : m_busyUntil = -1 : m_busyRest = "" : m_busyQueue = False : m_queued = ""
		CurrentRow = 24 : CurrentColumn = Len(m_prompt) + 1
		Set m_out = CreateObject("Scripting.Dictionary")
	End Sub

	Public Sub UseIOS()
		m_ios = True
		m_prompt = "Cat8kv#"
		CurrentColumn = Len(m_prompt) + 1
	End Sub

	' everything the device sends is queued for ReadString and shown on the screen
	Private Sub Out(s)
		m_pending = m_pending & s
		m_shown = m_shown & s
	End Sub

	' last Rows lines of what was shown, CRLF between rows (like the real Get2)
	Public Function Get2(r1, c1, r2, c2)
		Dim v, i, s, n
		v = Split(Replace(Replace(m_shown, vbCrLf, vbLf), vbCr, vbLf), vbLf)
		n = UBound(v) - Rows + 1
		If n < 0 Then n = 0
		s = ""
		For i = n To UBound(v)
			s = s & Left(v(i) & Space(Columns), Columns)
			If i < UBound(v) Then s = s & vbCrLf
		Next
		Get2 = s
	End Function

	Public Sub SetOutput(strCmd, strText)
		m_out(strCmd) = strText
	End Sub

	' time passes: a busy command that is due finishes, then any queued input runs
	Private Sub Tick(nSec)
		Dim v, i
		m_now = m_now + nSec
		If m_busyUntil < 0 Or m_now < m_busyUntil Then Exit Sub
		m_busyUntil = -1
		Out m_busyRest
		If m_queued <> "" Then
			v = Split(m_queued, vbCr)
			m_queued = ""
			For i = 0 To UBound(v) - 1
				Feed v(i) & vbCr
			Next
		End If
	End Sub

	Public Sub Send(s)
		SentLog = SentLog & s
		Feed s
	End Sub

	Private Sub Feed(s)
		If m_busyUntil >= 0 Then
			If m_busyQueue Then m_queued = m_queued & s
			Exit Sub
		End If
		If IsArray(m_left) Then
			' paging: Space = next page, Enter = next line, anything else = quit
			If m_ios Then
				Out String(9, Chr(8)) & Space(9) & String(9, Chr(8))
			Else
				Out vbCr & Space(40) & vbCr
			End If
			If s = " " Then
				EmitPage m_screenLen - 1
			ElseIf s = vbCr Then
				EmitPage 1
			Else
				m_left = Empty
				Out m_prompt
			End If
			Exit Sub
		End If
		If Right(s, 1) <> vbCr Then Exit Sub
		Dim strCmd, re, m
		strCmd = Left(s, Len(s) - 1)
		If strCmd = "" Then
			Out vbCrLf & m_prompt
			Exit Sub
		End If
		Out strCmd & vbCrLf
		Set re = New RegExp
		re.Pattern = "^set cli screen-length (\d+)$"
		If re.Test(strCmd) Then
			m_screenLen = CLng(re.Execute(strCmd)(0).SubMatches(0))
			Emit Array("Screen length set to " & m_screenLen)
			Exit Sub
		End If
		re.Pattern = "^terminal length (\d+)$"
		If re.Test(strCmd) Then
			m_screenLen = CLng(re.Execute(strCmd)(0).SubMatches(0))
			Emit Array()
			Exit Sub
		End If
		re.Pattern = "^(busyq?) (\d+)$"
		If re.Test(strCmd) Or strCmd = "dead" Then
			Out "Type escape sequence to abort." & vbCrLf & _
				"Sending 5, 100-byte ICMP Echos to 192.0.2.1, timeout is 2 seconds:" & vbCrLf & "."
			m_busyRest = "...." & vbCrLf & "Success rate is 0 percent (0/5)" & vbCrLf & m_prompt
			If strCmd = "dead" Then
				m_busyUntil = 1E9 : m_busyQueue = False
			Else
				Set m = re.Execute(strCmd)(0)
				m_busyUntil = m_now + CLng(m.SubMatches(1))
				m_busyQueue = (m.SubMatches(0) = "busyq")
			End If
		ElseIf strCmd = "hang" Then
			Out "working..." & vbCrLf
		ElseIf strCmd = "hangbig" Then
			Dim k
			For k = 1 To 30 : Out "line " & k & vbCrLf : Next
		ElseIf m_out.Exists(strCmd) Then
			Emit Split(m_out(strCmd), vbCrLf)
		Else
			Emit Array("                    ^", "syntax error, expecting <command>.")
		End If
	End Sub

	Private Sub Emit(vLines)
		m_left = vLines
		m_total = UBound(vLines) + 1
		If m_screenLen > 1 And m_total > m_screenLen - 1 Then
			EmitPage m_screenLen - 1
		Else
			EmitPage m_total
		End If
	End Sub

	Private Sub EmitPage(n)
		Dim i, nLeft, vRest()
		nLeft = UBound(m_left) + 1
		If n > nLeft Then n = nLeft
		For i = 0 To n - 1
			Out m_left(i) & vbCrLf
		Next
		If n = nLeft Then
			m_left = Empty
			Out m_prompt
			Exit Sub
		End If
		ReDim vRest(nLeft - n - 1)
		For i = n To nLeft - 1
			vRest(i - n) = m_left(i)
		Next
		If m_ios Then
			Out " --More-- "
		ElseIf nLeft = m_total Then
			Out "---(more)---"
		Else
			Out "---(more " & Int(100 * (m_total - nLeft + n) / m_total) & "%)---"
		End If
		m_left = vRest
	End Sub

	Public Function ReadString(vList, nTimeout)
		Dim i, p, best, bestIdx
		ReadCount = ReadCount + 1
		LastTimeout = nTimeout
		If ReadCount > 500 Then Err.Raise vbObjectError + 1, "FakeScreen", "runaway: 500 reads"
		Do
			best = 0 : bestIdx = 0
			For i = 0 To UBound(vList)
				If vList(i) = "" Then p = 1 Else p = InStr(m_pending, vList(i))
				If p > 0 And (best = 0 Or p < best) Then best = p : bestIdx = i
			Next
			' nothing yet: wait for a busy command that finishes within this read's timeout
			If best > 0 Or m_busyUntil < 0 Or m_busyUntil - m_now > nTimeout Then Exit Do
			Tick m_busyUntil - m_now
		Loop
		If best = 0 Then
			Tick nTimeout
			MatchIndex = 0
			ReadString = ""
			m_pending = ""
		Else
			MatchIndex = bestIdx + 1
			ReadString = Left(m_pending, best - 1)
			m_pending = Mid(m_pending, best + Len(vList(bestIdx)))
			If Not Synchronous Then
				p = InStr(m_pending, vbLf)
				If p > 0 Then m_pending = Mid(m_pending, p + 1) Else m_pending = ""
			End If
		End If
	End Function

	Public Function WaitForCursor(n)
		Tick n
		m_pending = ""
		WaitForCursor = False
	End Function

	Public Function [Get](r1, c1, r2, c2)
		[Get] = m_prompt
	End Function
End Class

Class FakeCrt
	Public Screen
	Private Sub Class_Initialize()
		Set Screen = New FakeScreen
	End Sub
End Class
