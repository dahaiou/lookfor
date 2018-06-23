
'! blawber
'! ssend ("help")

testnum = "0.1"
say "Running Test number " & testnum
ssend ("setmem 0600 8A8A8A8A")
'! say "sending: showmem 0600 10"
ssend ("showmem 0600 10")
'! say "FoundLine: """ & FoundLine & """"

expected="0600 | 8A 8A 8A 8A"

If InStr(FoundLine, expected) = 1 Then
  say "Test " & testnum & " OK."
Else 
	say "Test " & testnum & " FAILED."
	say "Expected: " & expected
	say "Received: " & FoundLine
End If
