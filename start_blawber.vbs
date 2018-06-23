' start_blawber.vbs
' call from .vbst routines to start blawber.exe as the slave app to test against
' ==============================================================================
if Blawber_started = False Then
	
	spawnnew

    SlaveExec.StdIn.WriteLine("..\Debug\Blawber.exe")
    SlaveExec.StdIn.WriteLine("/his dis")
    SlaveExec.StdIn.WriteLine("/prompt =:)>>")
    '! SlaveExec.StdIn.WriteLine("/echo )>>")
	SlaveFname = "Blawber.exe"
	
	SlavePrompt = "=:)>>"
	SlavePrompt = ")>>"
	FoundLine = lookfor (SlavePrompt)
	say FoundLine
	FoundLine = lookfor (SlavePrompt)
	say FoundLine
End If
