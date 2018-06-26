'! /<path>/VBScript CScript
'!_Language    : VBScript
'!_File        : lookfor.vbs
Dim ProgName, Version, ProgPackages, ProgNameString  '! Identify Program and loaded Modules
ProgName = "lookfor"
'!>_Author      : Henrik Soderstrom*
'!_Copyright   : (C) 2018, Henrik Soderstrom*
'!_License     : GPL2* (except for various snippets as noted in comments)
'! before running release, put current date in next line. Afterwards put back the mm-dd:s '#### <DEBUGLINE>
'!_Date        : 2018-06-dd
'! after running release, put current date in next line.                   '#### <DEBUGLINE>
'!_Date        : 2018-06-21-                                               '#### <DEBUGLINE>
'! before running release, remove the "p" in next two lines.               '#### <DEBUGLINE>
'! after running release, put the "p" back and step revision.              '#### <DEBUGLINE>
Version  = "0.06p_"				'_Version
'!_Version     : 0.06p_                          #### <DEBUGLINE>
'!_Description : Automated running/testing of win/dos console apps 
'!               (similar to expect, but in VBScript rather than tcl)
'!               - Start console app and send command lines to it (stdin)
'!               - "lookfor" and analyse the output coming back (stdout)
'!               - Run automated test scripts (extension .vbst)
'!               - Log results to logfile (.vbstlog)
'! ====================================================================================================
'_* Copyright, authorship and license apply as stated, EXCEPT for:
'   various snippets from other sources, noted in comments in each case
'_Usage      :: See the help_message function below

' TODO: Organise the references to "other sources" better, eg. the following bit:
' Based partly on: "A simple interactive VBScript shell."
' @see http://www.kryogenix.org/days/2004/04/01/interactivevbscript
'
' ====================================================================================================
' (NOTE: rulers of exactly 100 = equal signs, preceded by quote ' and whitespace as necessary)
'_History:
' V0.02 Saved 2018-06-01 with Package:inc.vbs_V0.02-01
' V0.03 Started 2018-06-01 with Package:inc.vbs_V0.03-01
' V0.03 Saved/Released 2018-06-06 with Package:inc.vbs_V0.03-01
' ====================================================================================================
' V0.04 Started 2018-06-06 with Package:inc.vbs_V0.04-01
'		initial stuff:
' 		- removed debugging stuff(at end of inc1.vbs):
' 		- dim trexregex, sub setrg (s), Sub trex(ByRef cmdline),
'		- Sub trx(ByRef cmdline), Sub trx_old(ByRef cmdline)
' V0.04 Release Notes 2018-06-13
'		Lots of new stuff added, mainly in inc1.vbs
'		This list highlights the main things, probably not complete:
'		o Running TestCases from .vbst files is working now in a
'		  .bare-bones sort of way, still missing lots, such as:
'		  .output to logfiles, bookkeeping, statistics etc.
'		o Started adding some structured formatting markup in comments
'		  .(eg: '_h1 Misc funtions) with some loose ideas of
'		  .automatic documentation generation a la perl autodoc
'		  .TODO: Google for existing solutions as alternative
'		o "Dot notation" to simplify input without needing to type quotes
'		  .instead of say "hello" you can go .say hello
'		  .dot notation lines are also preprocessed, see next point.
'		o Preprocessing of command lines with substitution of expressions:
'		  . eg. ".say Running Test: {TCnum}" substitutes value of TCnum
'		  . or ".say 3*4 gives: {3*4}" substitutes 3*4 with 12
'		o Unix-style option handling eg. cmd -abc --long-opt1 arg1 arg2
'		  .Main routine: getopts (optstr, ByRef cmdline, ByRef opts_found)
'		  .Easy to implement in any routine, which can then be called
'		  .with short or long options either from the command prompt
'		  .or from other routines
'		  .
'		o Improved startup of slave apps for testing
'		  .slave app can be started from a cmd.exe prompt or directly
'		  .redirecting of stderr only seems to work from cmd.exe
'		  .lookfor can now start a slave instance of itself
'		  .and run test scripts on itself
'		  .the getopts routine above, having a suitably low but
'		  .non-trivial complexity level, is actually becoming
'		  .the pilot test object used in the development of the
'		  .TestCase functionality
'		o Handling command line arguments (on invocation)
'		  .lookfor.vbs now checks for args passed when invoked
'		  .eg. from a dos command line. lookfor's command prompt
'		  .can be set in this way, which is useful when starting
'		  .lookfor as a slave instance for testing.
'		o Various small utility functions, eg:
'		  . sayvar("variable") - print out "variable=<value>"
'		  . good for debugging. Major drawback: var MUST be global
'		  . would be perfect to implement as a preprocessing macro
'		  . sayvarq - as sayvar, with single-quoted '<value>'
'		  . so any additional whitespace is easily visible
'		  . sub test_rx(s,sPat) to test regex functionality
'		  . Sub ListProcessRunning() - show processes, copied
'		  . from stackoverflow. to be removed later
'		  .
'		o sDB... "silly DataBase" routines 
'		  .(sDBadd, sDBdelete, sdBcheck, sDBreset)
'		  .Store keywords in a string variable, separated by |
'		  .similar to a PATH variable
'		o Selective debugging with the saydbg command
'		  .saydbg "@key1 anytext" can be sprinkled around the code
'		  .to help with troubleshooting and debugging
'		  .Whether saydbg prints anything or not is controlled
'		  .selectively with the @key value and some global variables:
'		  .DBG_enabled, DBG_current, DBG_banner
'		  .
'		o Easter Egg: The 100 Hundred Doors problem included
'		  at end of inc1.vbs, now working up to n doors.
'		  HURRY and check it out: probably to be removed soon 
' 
' V0.05 Started 2018-06-13 with Package:inc.vbs_V0.05-01
' V0.05 Release Notes 2018-06-21
'		This list highlights the main things, probably not complete:
'		o Running TestCases from .vbst files works in a
'		  .bare-bones way, still missing lots, such as:
'		  .output to logfiles, bookkeeping, statistics etc.
'		o Started adding formatting markup for vbsdoc
'		  for automatic doc generation
'		o Started modifying vbsdoc to my needs
'		o Some improvemed debugging functions eg. 
'		  . vbsCodetoSayLocalVar and vbsCodetoSayLocalVarq
'		o ssend now sticks a "slave marker" (ie a ">") at the
'		  beginning of every line coming from slave stdout
'		o Easter Egg: The 100 Hundred Doors problem still included
'		  and was adapted a little.
'		 
' V0.06 Started 2018-06-21 with Package:inc.vbs_V0.06-01
'		 
'		 
'		 
'		 
'		 
'!@TODO:
' o Standardised detection of whether ar slave app is active or not
' o 
' o mechanisms to change the child process prompt with just one command
'   ie. with one command change both: a.) Child's prompt string, by command to child
'   AND b.) prompt pattern used by Parent when reading from slave's stdout
'		 
'		 
'		 
' ====================================================================================================
' ====================================================================================================

'
' Option Explicit

Dim HideSlavePrompt

ProgPackages = ""
ProgNameString = ProgName & " V" & Version
HideSlavePrompt = True

DIM GcmdLine  '! Command line
GcmdLine = ""
Dim Gtest1, Gtest2  '! for testing

Gtest1 = "Global test 1"
Gtest2 = "Global test 2"


Dim FoundLine '! Last line read back from stdout of Slave process
Foundline = ""
Set SlaveShell = Nothing
Set SlaveExec  = Nothing


' TODO: "prompt" is used everywhere below to mean slaveprompt. Needs to be
'		 refactored as SlavePrompt or similar and then myprompt to prompt

Dim MyPrompt
MyPrompt = ProgName & ":> "
SlaveOutFlag = ""
SlavePrompt = ""
SlaveFname = ""						' SlaveFname holds filename of slave app when active
SlaveCmdFn = ""						' SlaveCmdFn holds filename of dos shell, normally "cmd.exe", when active

Dim ErrArray

Class ErrCopy
	Public Number, Description, Source
	Sub Reset
		Number 		= 0
		Description = ""
		Source		= ""
	End Sub
End Class
Set MyErr = New ErrCopy
MyErr.Reset

getArgs

'! Get command-line arguments (from the DOS/Win command line)
'! =====================================================================================================
Sub getArgs()
	For each arg in wscript.arguments
		If InStr (arg, "-P") = 1 Then
			p = Mid(arg, InStr(arg, "/") + 1)
			'say "p="&p
			If InStrRev(p, "/") > 1 Then
				p = Left(p, InStrRev(p, "/") - 1)
				'say "p="&p
				MyPrompt = p
			End If
		ElseIf InStr (arg, "-R") = 1 Then
			p = Mid(arg, 3)
			If len (p) > 1 Then MyPrompt = p
		End If
	Next
End Sub 


Sub SetMyPrompt(p)
	myprompt = p
End Sub


' TODO: Is this delete needed here ???
' Delete LF_InitScript
Private Const LF_InitScript    = "%USERPROFILE%\init.vbs"

' WScript.StdOut.Write("lookfor Bk 1:> ")

'! Sub spawnnew(ByVal filename)
'! Spawn a new cmd.exe process to send input to and receive output from
Sub spawnnew()
	Set SlaveShell = Nothing
	Set SlaveExec  = Nothing

	Set SlaveShell = WScript.CreateObject("WScript.Shell")
	Set SlaveExec = SlaveShell.Exec("cmd.exe 2>&1")
	SlaveCmdFn = "cmd.exe"
	SlaveFname = ""						' No slave app started yet
    
	SlaveExec.StdIn.WriteLine("prompt=$P$G(): ")
    SlavePrompt   = ">(): "			'! set SlavePrompt to deal with a WIN/DOS Console
	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine

End Sub '! 'Sub spawnnew(ByVal filename)

'! Spawn a New Slave Process
'!
'! @param  cmdline  The command line used (in a DOS/Win prompt) to start the new process
'! @param  prompt   The prompt string to be passed to the slave process on initiation
'! 					The slave must be able to set its command line prompt according to this.
'!
Sub NewSlve (cmdline, prompt)

	Set SlaveShell = Nothing
	Set SlaveExec  = Nothing

	Set SlaveShell = WScript.CreateObject("WScript.Shell")
	Set SlaveExec = SlaveShell.Exec(cmdline)

    SlaveFname = cmdline

	SlavePrompt = prompt
	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine
	'FoundLine = SlveReadUpto (SlavePrompt)
	'say FoundLine
End Sub ' Sub SpawnCliCmd (cmdline, prompt, InitialCommands)

Sub Spawnlookfor ()

	HideSlavePrompt = False
	SlaveOutFlag = ">"
	' NewSlve "cscript.exe //NoLogo ./lookfor.vbs -P""/<child>:/""", "<child>:"
	NewSlve "cmd.exe /c cscript.exe 2>&1 /NoLogo lookfor.vbs -P""/<child>:/""", "<child>:"

End Sub

Sub NewCmdLookfor ()
	Set SlaveShell = Nothing
	Set SlaveExec  = Nothing

	Set SlaveShell = WScript.CreateObject("WScript.Shell")
	Set SlaveExec = SlaveShell.Exec("cmd.exe /c cscript.exe 2>&1 /NoLogo lookfor.vbs -P""/<child>:/""")
	SlaveCmdFn = "cmd.exe"
	SlaveFname = "lookfor.vbs"
    
	HideSlavePrompt = False
	SlavePrompt = "<child>:"

	SlaveOutFlag = ">"

	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine
	'FoundLine = SlveReadUpto (SlavePrompt)
	'say FoundLine

End Sub

Sub blawber ()
	
	spawnnew

    SlaveExec.StdIn.WriteLine("..\Debug\Blawber.exe")
    SlaveExec.StdIn.WriteLine("/his dis")
    SlaveExec.StdIn.WriteLine("/prompt =:)>>")
    '! SlaveExec.StdIn.WriteLine("/echo )>>")
	
	SlaveFname = "Blawber.exe"

	SlavePrompt = "=:)>>"
	'SlavePrompt = ")>>"
	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine
	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine
End Sub ' Sub blawber ()


Sub send (cmdline)
	if SlaveExec Is Nothing then
		sayerr "Error: No slave app to send to."
		Exit Sub
	End If
    SlaveExec.StdIn.WriteLine(cmdline)
End Sub ' send(cmdline)

Sub old_ssend (ByRef cmdline)
	Dim opts_found		' Make sure opts_found is local
	opts_found = ""
	saydbg "calling getopts with :el " & cmdline
	getopts ":el", cmdline, opts_found

	if SlaveExec Is Nothing then
		sayerr "Error: No slave app to send to."
		Exit Sub
	End If
	saydbg "calling find_opt ( l, " & opts_found
	if find_opt("l", opts_found) Then
		TClog "Sending: " & cmdline
	End If
    TClog "Sending: " & cmdline
	SlaveExec.StdIn.WriteLine(cmdline)
	FoundLine = SlveReadUpto (SlavePrompt)
	If find_opt("e", opts_found) And len(FoundLine) > 0 _
		then say SlaveOutFlag & Replace (FoundLine, VBCrLf, SlaveOutFlag & VBCrLf)

End Sub ' ssend(cmdline)

Sub rsend (cmdline)
	if SlaveExec Is Nothing then
		sayerr "Error: No slave app to send to."
		Exit Sub
	End If
    SlaveExec.StdIn.WriteLine(cmdline & " 2>&1")
	FoundLine = SlveReadUpto (SlavePrompt)
	say FoundLine
End Sub ' send(cmdline)


' ==============================================================================
' TODO: When tested, move routines from inc1.vbs in here
' ==============================================================================
import "inc1.vbs"

' ==============================================================================

Sub saynext ()
	FoundLine = SlveReadUpto(SlavePrompt)
	say FoundLine
End Sub

'! Print a text on the console
'! @param  s    The text to be printed
Sub say (s)
	wscript.echo s
end sub

sub sayerr (s)
	WScript.StdErr.WriteLine s
end sub

'Marker: saydbg was here

sub sayq (s)  ' say quoted ie. within quotes
	WScript.StdErr.WriteLine "'" & s & "'"
end sub


'! Import initialization script if present.
Private Sub ImportLF_InitScript()
	Dim sh, fso, path, initScriptExists

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	path = sh.ExpandEnvironmentStrings(LF_InitScript)
	initScriptExists = fso.FileExists(path)

	Set sh  = Nothing
	Set fso = Nothing

	If initScriptExists Then Import path
End Sub

'! Print usage information.
Private Sub Usage()
	WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
		& vbTab & "help                      Print this help." & vbNewLine _
		& vbTab & "?                         Open the VBScript documentation." & vbNewLine _
		& vbTab & "? ""keyword""               Look up ""keyword"" in the documentation." & vbNewLine _
		& vbTab & "                          The helpfile (" & Documentation & ") must be installed" & vbNewLine _
		& vbTab & "                          in either the Windows help directory, %PATH%" & vbNewLine _
		& vbTab & "                          or the current working directory." & vbNewLine _
		& vbTab & "import ""\PATH\TO\my.vbs""  Load and execute the contents of the script." & vbNewLine _
		& vbTab & "exit                      Exit the shell." & vbNewLine _
		& vbNewLine & "Customize with an (optional) init script '" & LF_InitScript & "'." & vbNewLine _
		& vbNewLine
End Sub

'! Import the first occurrence of the given filename from the working directory
'! or any directory in the %PATH%.
'!
'! @param  filename   Name of the file to import.
'!
'! @see http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/
Private Sub Import(ByVal filename)
	Dim fso, sh, file, code, dir

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	if InStr(filename, ".") = 0 then filename = filename & ".vbs"
	shortfilename = filename
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		' filename is not absolute
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			' file doesn't exist in the working directory => iterate over the
			' directories in the %PATH% and take the first occurrence
			' if no occurrence is found => use filename as-is, which will result
			' in an error when trying to open the file
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If

	Set file = fso.OpenTextFile(filename, 1, False)
	code = file.ReadAll
	file.Close

	myErr.Reset
	ExecuteGlobal(code)

	If Err.Number <> 0 then
		WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
		myErr.Number = Err.Number
		myErr.Description = Err.Description
		myErr.Source = Err.Source
		ErrArray = Array (Err.Number, Err.Description, Err.Source)
	End If

	Set fso = Nothing
	Set sh = Nothing
End Sub '! Private Sub Import(ByVal filename)

'! Run a vbsh command
'! NOTE: if interacting by sending commands to a slave app,
'! the returned output from the most recent command is found in variable FoundLine
Sub RunVbshLine (line)

		CmdTokens = split(Trim(line))
		cmd = ""
		if ubound(CmdTokens) >=0 Then cmd = CmdTokens(0)

		' Alias preprocessing
		'If cmd = "dbg" Then line = "." & line		' EXPERIMENTAL ...   **** DEBUG

		If line = "q" Then
			If SlaveExec Is Nothing then ' Do nothing
			Else
			    SlaveExec.StdIn.WriteLine("exit")
				Set SlaveShell = Nothing
				Set SlaveExec  = Nothing
			end if
			Wscript.Quit 

		ElseIf line = "help" Then
			help
		ElseIf line = "h" Then
			help
		ElseIf line = "?" Then
			help
		
		'! NOTE: Once Slave app is started you can send commands to it directly
		'! 		 by simply preceding with ">" or "_" eg. _/echo "hello"
		
		ElseIf Left(line, 1) = ":" Then
			say "Test output(preprocess_cmdline): " & preprocess_cmdline (mid(line,2))
		ElseIf Left(line, 1) = "." Then
			'RunVbshLine preprocess_cmdline (mid(line,2))
			myExecuteGlobal preprocess_cmdline (mid(line,2))
		ElseIf Left(line, 1) = ">" Then
			ssend (mid(line,2))
		ElseIf Left(line, 1) = "_" Then
'		    saydbg "calling ssend with line:"&mid(line,2)
			ssend (mid(line,2))
		ElseIf True Then
			myExecuteGlobal line
		Else ' This Else is never reached	
			On Error Resume Next
			Err.Clear
			ExecuteGlobal line
			If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
		End If

End Sub '! Sub RunVbshLine (line)

Sub myExecuteGlobal (line)
	GcmdLine = line
	On Error Resume Next
	Err.Clear
	myErr.Reset
	
	ExecuteGlobal line 
	Err_Number = Err.Number
	If Err.Number <> 0 then
		WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
		myErr.Number = Err.Number
		myErr.Description = Err.Description
		myErr.Source = Err.Source
		ErrArray = Array (Err.Number, Err.Description, Err.Source)
	End If

	On Error Goto 0
End Sub ' Sub myExecuteGlobal


'! Run sub Main:
Main

Sub Main()
	Dim line

	'ImportLF_InitScript
	'Usage

	say "Welcome to " & ProgNameString & " (h or ? for help, q to quit)"


	Do While True
		WScript.StdOut.Write(MyPrompt)

		line = Trim(WScript.StdIn.ReadLine)
		Do While Right(line, 2) = " _" Or line = "_"
			line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
		Loop

		Do While Right(line, 3) = " ++" Or line = "++"
			line = RTrim(Left(line, Len(line)-2)) & """" & " & VBCrLf & """ & Trim(WScript.StdIn.ReadLine)
		Loop

		If LCase(line) = "exit" Then Exit Do
		
		RunVbshLine line

	Loop
End Sub '! Sub Main()

