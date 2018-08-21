'! /<path>/VBScript CScript
'!_Language    : VBScript
'!_File        : lookfor.vbs
Dim ProgName, Version, ProgPackages, ProgNameString  '! Identify Program and loaded Modules
ProgName = "lookfor"
'!_Author      : Henrik Soderstrom*
'!_Copyright   : (C) 2018, Henrik Soderstrom*
'!_License     : GPL2* (except for various snippets as noted in comments)
'  before running release, put current date in next line. Afterwards put back the mm-dd:s '#### <DEBUGLINE>
'!_Date        : 2018-mm-dd
'  after running release, put current date in next line.                   '#### <DEBUGLINE>
'!_Date        : 2018-07-22-                                               '#### <DEBUGLINE>
'  before running release, remove the "p_" in next two lines.              '#### <DEBUGLINE>
'  after running release, put the "p" back and step revision.              '#### <DEBUGLINE>
Version  = "0.07p02_"				'_Version
'!_Version     : 0.07p_                          #### <DEBUGLINE>
'!_Description : Automated running/testing of CLI (Command Line Interface) applications
'!				 (aka "console apps") in the win/dos environment
'!               Inspired by expect to some extent, but implemented in VBScript rather than tcl
'!               - Start console app and send command lines to it (stdin)
'!					(standard input)
'!               - "lookfor" and analyse the output coming back (stdout)
'!               - Run automated test scripts (extension .vbst)
'!               - Log results to logfile (.vbstlog)
'! ====================================================================================================
'_* Copyright, authorship and license apply as stated, EXCEPT for: 
'   various snippets from other sources, noted in comments in each case
'_Usage      :: See the help_message function below

'!@TODO: Organise the references to "other sources" better, eg. the following bit:
'! Based partly on: "A simple interactive VBScript shell."
'! @see : A simple interactive VBScript shell. http://www.kryogenix.org/days/2004/04/01/interactivevbscript
'
' ====================================================================================================
' (NOTE: rulers of exactly 100 = equal signs, preceded by quote ' and whitespace as necessary)
'_History:
' V0.02 Saved 2018-06-01 with Package:inc.vbs_V0.02-01
' V0.03 Started 2018-06-01 with Package:inc.vbs_V0.03-01
' V0.03 Saved/Released 2018-06-06 with Package:inc.vbs_V0.03-01
' V0.04 Released 2028-06-13 with Package:inc.vbs_V0.04-01
' V0.05 Released 2028-06-21 with Package:inc.vbs_V0.05-01
' V0.06 Released 2028-07-22 with Package:inc.vbs_V0.06-01
'		(also including TestCase.vbst V0.06-01, misc.vbst V0.06-01)
' ====================================================================================================
' V0.04 Release Notes 2018-06-13: Removed, see earlier releases
' V0.05 Release Notes 2018-06-21: Removed, see earlier releases
'		 
' V0.06 Started 2018-06-21 with Package:inc.vbs_V0.06-01
' 		o During development _Preliminary_ V0.06 is reflected as Version="0.06p_"
' V0.06 Release Notes 2018-07-22
'		This list highlights some main things, probably not complete:
'	_h2 About these Release Notes
'		o Release Notes and comments from earlier versions (0.04-05) removed but
'		  . many of the same points are commented on here.
'	_h2 Doc Formatting: Handling doc markup included in the code:
'		o Googled around and settled, at least for now, on Ansgar Wiecher's vbsdoc for this
'		  and started modifying it:
'		  - Put linefeeds back in for some cases. vbsdoc removes them consistently
'		  - A lot more to do to make it do what I want
'	_h2 Importing and including files
'		o The main program lookfor.vbs now imports several .vbs/t files on startup:
'		  - inc1.vbs 		- Several help routines, plus imports the other ones
'		  - TestCase.vbst	- TestCase routines (plus some other stuff creeping in)
'		  - Misc.vbst		- Miscellaneous
'		  The last two, with the .vbst extension use the same preprocessing 
'		  that is available to TestCase files.
'		o Note: all these files can be reimported into a live session by ".import inc1",
'		  . useful in debugging. Or also reimported individually by ".rtf filename" 
'		o In principle we try to avoid resetting global vars on reimport,
'		  . eg. debug settings are preserved. That said, a few have probably been missed.
'		o Mechanism implemented to skip parts of code on reimport, eg. for
'		  . class definitions as classes cannot be redefined
'		o Implementation: A global flag (eg. G_MyClass_defined) is set on initial import 
'		  . Then REimport can be detected and those blocks of code skipped by setting
'		  . another global flag: GlobalDiscardNextBlock = True
'		o TODO: This skipping mechanism needs polishing up a bit
'		o The main program _can_ be reimported live with ".import lookfor"
'		  . but many global vars will be clobbered by resetting to initial values
'		o TODO: Consider separate file(s) to reinitialise global variables
'		o TODO: Make the import/include more general and change to .vbsx (extended vbs)
'		o TODO: Make the main lookfor.vbs fairly bare-bones and most functionality
'		  . implemented in imported .vbsx modules
'		o imported modules can announce their presence and version by setting global vars
'	:h2 Running TestCases from .vbst files:
'		o TestCase files are run with ".rtf filename" (the default extension is .vbst )
'		o Consequently, most the comments above about "Importing and including files"
'			also applies to running testcases.
'		o Even though this functionality was the intended main purpose of the lookfor program,
'		  the progress in this area has been rather modest.
'		  A lot of time slipped by deciding how to implement logging to files
'		  . and global housekeeping of test results, and these questions are
'		  . still up in the air quite a bit.
'		o output to a logfiles implemented as a simple mechanism for now. Not sure if
'			it needs to be further developed.
'		  . still missing bookkeeping, statistics, different levels of logging detail etc.
'		  . TODO: Considering defining TC classes rather than global variables
'	:h2. Major revamp in progress, of the preprocessing routines used both
'		in direct command input and during import/include of .vbst files
'		 o Regexp'es have mostly replaced the pure vbs-string-function approach, and var/expression
'			 substitution is enhanced, including local vars and shell environment vars
'		 o The new functionality can already be tested interactively by "comma-notation"
'			(instead of dot) eg. instead of ".say hello {name}" you can go ",say hello $name"
'			Note that the comma-notation is only a temporary solution to help with debugging.
'			The new routines, once tested, will replace the old as standard, and be used
'			through the normal dot-notation.
'		 o The "old" routines, to be replaced by the new, but still being used for now (file: inc.vbs):
'			- Sub RunTestFile(ByVal filename)			- run or import a .vbst file
'			- Function preprocess_cmdline (cmdline)		- Main routine to preprocess a command line, calls replace_args
'			- Function replace_args (argline)			- Handles the actual substitution of vars and expressions
'			- 
'		 o The new routines and elements being introduced (most of these in file: TestCase.vbst):
'			- Function findFileName(ByVal filename) *
'				- This is a separate routine, duplicating the file-finding logic existing in RunTestFile
'			- Sub RunTest(ByVal filename)		- New routine to run test cases
'			- Function ssubst (s)		- Experimental routine, not used, to be removed later
'			- Function cmdsubst (s)		- Main routine to preprocess a command line, calls argsubst(s)
'			- Function argsubst(s)		- This is where the actual substitution is done
'			- Function checkQ (s)	 *	- Removes comments, quotes and curly-brace-enclosed elements
'			- Function unComment (s) *	- Removes trailing comments
'				* these are in file inc.vbs, Note: checkQ and unComment do NOT use RegExp
'		 o Various new RegExp defined to handle the parsing/preprocessing of input lines:
'			- oRxIniPunct		- Find initial punctuation character
'			- oRxIniToken		- Find first token in line
'			- oRxSplitmark		- Find next one of the punctuation characters used in preprocessing
'			- oRxSplitDquote	- Find closing double-quote
'			- oRxRightCurly		- Find closing curly-brace
'			- oRxCurlyMissing	- Match case where closing curly-brace is missing
'			- oRxCurlyEmpty		- Match whitespace or empty string within curly-braces
'			- oRxCurlyNormal	- Match text within curly-braces
'		 o The "old" runTestFile routine has been enhanced by handling curly-braces "{}"
'			as start and stop tokens for multi-line code blocks. The curly-brace notation
'			is meant to replace the old tokens "<:" and ":>". The old ones still work for now.
'		 o New syntax for variable and expression substitution:
'			The "new" routines allow substitution of "name" to be expressed in different ways:
'			- $name substitution is deferred to execution time and happens in local scope.
'				Being able to substitute local vars is new and very useful
'			- ${name} equivalent to $name for simple var substitution, however:
'				-- must use this one if you need to separate the var from characters following it.
'					eg. ${name}hello works, where $namehello refers to a different variable
'				-- (not working yet) can include expressions eg. ${linecount - v2 + 4}
'					Though unclear whether this is useful or not
'			- %name substitutes variable name from the shell environment eg. %path
'			- %name% or %{name} are equivalent to %name, sometimes necessary
'				to separate the var from characters following it.
'			- {name} this is the "old" syntax used up to now, and it still works, in parallell
'				with the ones above. Some advantages "+ plus" and disadvantages "- minus":
'				+ can substitute expressions eg. {4+5} gives 9
'				+ can even run code eg. ,say {4: say "Way out of my depth": runtest "opttest1"}
'				  (the 4 is necessary since the first statement must work as a variable assignment)
'				  Not clear if this is a plus or a minus. Certainly dangerous and NOT recommended
'				- A big minus: substitution happens in global scope. Local vars are not visible
'				- Substitution is at parse-time, substituted values may end up being out of date.
':h1.Todo
'!@TODO: Todo and ideas
'		o Improved startup of slave apps for testing:
'			- No progress on this. medium to low priority. See V0.04 Release Notes 2018-06-13
'		o Testing of the getopts routine:
'			- Very modest progress. medium priority but is important as development help for
'				the testcase functionality
'		0 Debugging routines: sayvarq, sayerrq
'			- Should be easy to do, and valuable, now that local vars work in the new routines
'		o "Extended vbs" preprocessing.
'			This is the same, but a better name than "Importing and including files" used above
'			- On Error, identify offending multi-line block by start and end line numbers
'			- Option to show n lines of the offending block (currently all is shown)
'			- Option to log such error messages to file
'			- Error if block start token found before end: "nested multi-line blocks not allowed"
'			- Error for nested curly-brace elements
'			- Handle quoted elements within curly brace eg. {this text "includes quote with {} in it" etc.}
'			- getUntilExceptQuoted (s, c) - search string s until char c found, skipping any c included in quotes
'			- 
'			- 
'		o User Documentation
'			- It's high time to start creating some proper documentation.
'				All that exists now are these comments
'				The html generated by vbsdoc is useful as programming reference
'				but less practical for user documentation
'		 
'		o Standardized detection of whether a slave app is active or not
'		o Further development to handle structured formatting markup
'		  based on modifying Ansgar Wiecher's vbsdoc (at least for the moment)
'		o Mechanisms to change the child process prompt with just one command
'		  ie. with one command change both: a.) Child's prompt string, by command to child
'		  AND b.) prompt pattern used by Parent when reading from slave's stdout
'		 
' V0.07 Started 2018-07-22 with Package:inc.vbs_V0.07-01
'		(as well as TestCase.vbst V0.07-01, misc.vbst V0.07-01)
' 		o During development _Preliminary_ V0.06 is reflected as Version="0.06p_"
' V0.07p01 2018-07-27 with Packages:inc.vbs etc. at V0.07-01
'		 Preliminary Snapshot just BEFORE doing the following actions:
'		o Remove the old multi-block syntax, with start and stop markers "<:" and ":>"
'		o replace <: old notation :> with { new } curly-brace notation in all .vbst files
'		o Connect .dot notation to the new routines, and temporarily the old ones to ,comma notation
'		o Remove misc. other old stuff
'		NOTE: The mentioned changes have NOT YET been done here in V0.07p01.
'		They are done directly after saving this snapshot
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

Dim GtempVar  '! for temporary manipulation
GtempVar = ""

Dim FoundLine '! Last line read back from stdout of Slave process
Foundline = ""
Set SlaveShell = Nothing
Set SlaveExec  = Nothing

'! ====================================================================================================
'! ====================================================================================================
'! ====================================================================================================
'! ====================================================================================================
Set objShell = CreateObject("Wscript.Shell")
WScript.Echo objShell.CurrentDirectory
' objShell.CurrentDirectory = "G:\Dropbox\u\henk\PROG\VStudio\Projects\Blawber\lookfor\"
objShell.CurrentDirectory = Left (WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
Wscript.Echo objShell.CurrentDirectory
'! ====================================================================================================
'! ====================================================================================================
'! ====================================================================================================
'! ====================================================================================================


' TODO: Check that slaveprompt has been properly refactored as SlavePrompt, then
'		refactor the main prompt as prompt (or mainPrompt)

Dim MyPrompt
MyPrompt = ProgName & ":> "
SlaveOutFlag = ""
SlavePrompt = ""
SlaveFname = ""						' SlaveFname holds filename of slave app when active
SlaveCmdFn = ""						' SlaveCmdFn holds filename of dos shell, normally "cmd.exe", when active

Dim GlobalDiscardThisBlock, GlobalDiscardNextBlock
GlobalDiscardThisBlock = False
GlobalDiscardNextBlock = False

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

' Create Global Shell and FileSystem objects, always available from anywhere
Dim GoSh :Set GoSh  = CreateObject("WScript.Shell")
Dim GoFS :Set GoFS  = CreateObject("Scripting.FileSystemObject")

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

sub sayq (s)  ' say quoted ie. within quotes
	wscript.echo "'" & s & "'"
end sub

sub sayerr (s)
	WScript.StdErr.WriteLine s
end sub

sub sayerrq (s)
	WScript.StdErr.WriteLine "'" & s & "'"
end sub

'Marker: saydbg was here

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

		ElseIf line = "help" _
		Or line = "h" _
		Or line = "?" Then
			help
			Exit Sub
		End If
		
		'! NOTE: Once Slave app is started you can send commands to it directly
		'! 		 by simply preceding with ">" or "_" eg. _say "hello"
		
		If Left (line, 1) = "%" Then line = eval (funcsubst (line))

		If Left(line, 1) = ";" Then
			say "Test output(preprocess_cmdline): " & preprocess_cmdline (mid(line,2))
		ElseIf Left(line, 1) = "," Then
			'RunVbshLine preprocess_cmdline (mid(line,2))
			myExecuteGlobal preprocess_cmdline (mid(line,2))
		ElseIf Left(line, 1) = ":" Then
			say "Test output(cmdsubst): " & cmdsubst (line)
		ElseIf Left(line, 1) = "." Then
			myExecuteGlobal cmdsubst (line)
		ElseIf Left(line, 1) = ">" Then
			ssend (mid(line,2))
		ElseIf Left(line, 1) = "_" Then
			ssend (mid(line,2))
		Else
			myExecuteGlobal line
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
Main Wscript.Arguments

Sub Main(wargs)
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

