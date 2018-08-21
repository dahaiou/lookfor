
' Global Declarations and Initialisation
' =================================================================================================
IncludeName = "inc1.vbs"
IncludeVersion = "0.07-02"
inc1NameString = IncludeName & " V" & IncludeVersion
say "Including: " & inc1NameString

ProgPackages = ProgPackages & " " & Include_blurb

Dim g_opts_found			' #### DEBUG: global var holding options found from last call to getopts
g_opts_found=""

Dim oRx, oMatch			'! Global Regex, available to use anywhere
If vartype (oRx) = 0 Then	' Only init first time: Conserving variable values if this code is reimported
	Set oRx = New RegExp
	oRx.global = True
	oRx.ignorecase = False
	oRx.pattern = "^(\s*)"
	Set oMatch = oRx.Execute("Dummy text")
End If

' Escape-mark punctuation, complete oRxEscMark, reduced: oRxEscMark2
Dim oRxEscMark	: Set oRxEscMark	= NewRegExp("^([^\S\n]*)([',.;:_\\\|\@\#\~!$%&=?\+\-\*\/])([\s\S]*)$", True, True)	' Complete set
Dim oRxEscMark2	: Set oRxEscMark2	= NewRegExp("^([^\S\n]*)([',.;:_\#%])([\s\S]*)$", True, True)						' Reduced set

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
' Routines
' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0

'! Prepare a new regular expression.
'! By Ansgar Wiechers, See Copyrights Ref.1
'! Name changed from CompileRegExp to NewRegExp
'!
'! @param  pattern      The Regular Expression to match
'! @param  ignoreCase   True = Ignore case, False = Case-sensitive match
'! @param  searchGlobal True = Return all matches, False = Stop after first match
'! @return A new regular expression object.
Private Function NewRegExp(pattern, ignoreCase, searchGlobal)
	Set NewRegExp = New RegExp
	NewRegExp.Pattern    = pattern
	NewRegExp.IgnoreCase = Not Not ignoreCase
	NewRegExp.Global     = Not Not searchGlobal
End Function

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O


' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
'! Left-trim a string, using a RegEx
'! @param  s			The string to trim
'! @return The string with initial whitespace (except newline) removed
Private oRxLTrim : Set oRxLTrim = NewRegExp("^[^\S\n]*", True, True)
Function rxLTrim (s): 
	rxLTrim=oRxLTrim.replace(s,"")
End Function

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
'! Right-trim a string, using a RegEx
'! @param  s			The string to trim
'! @return The string with trailing whitespace (except newline) removed
Private oRxRTrim : Set oRxRTrim = NewRegExp("[^\S\n]*$", True, True)
Function rxRTrim (s)
	rxRTrim=oRxRTrim.replace(s,"")
End Function

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
'! Left- and right-trim a string, using a RegEx
'! @param  s			The string to trim
'! @return The string with leading and trailing whitespace (except newline) removed
Function rxTrim (s)
	rxTrim=rxRTrim(rxLTrim(s))
End Function


' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
Sub ssend (ByRef cmdline)
	'! Usage: ssend -E -l <line>
	'!@param -E - Echo off (on is default) Do not echo return text from slave app (only read into FoundLine)
	'!@param -v - Verbose on (off is default) Echo <line> just before sending it
	'!@param -l - Logging on (off is default) Echo <line> just before sending it
	' Ouch: There are a surprising number of options at play here:
	' 1. Echo or "report" the line which is about to be sent, before actually sending it
	'		1a. To console stdout: yes or no
	'		1b. To TClogfile if open: yes or no
	'			1b1. Warn if TClogfile not open: yes or no
	' 2. Write the line(s) read back from slave's stdout
	'		2a. To console stdout: yes or no
	'			2a1. Prepending slavemark character: yes or no
	'		2b. To TClogfile if open: yes or no
	'			2b1. Warn if TClogfile not open: yes or no
	'			2b2. Prepending slavemark character: yes or no
	' HARD to find the right terminology for these different alternatives
	'	1. Echo, Report, Chatty, Send-Echo, Echo-Send, Echo-Presend, EchoBeforeSend, EBefS
	'	2. Back-Echo, Slave-Echo, EchoSlaveOutput, Back-Echo, Slave-Echo
	'	
	'			Function:					Environment Variable			options	  Defaults
	'		1a. EchoSlaveInput, 			EnvVariable:	ESlvIn			-i, -I	: True, False (default)
	'		1b. TClogBeforeSend. 		 	EnvVariable:	LSlvIn			-s, -S *: True, False (default)
	'		2a. EchoSlaveOutput				EnvVariable:	ESlvOut			-o, -O	: True (default), False
	'			">" or "" or other			EnvVariable:	ESlvMark 				: ">" is default
	'		2b. TClogSlaveOutput, 			EnvVariable:	LSlvOut			-l, -L *: True (default), False
	'			">" or "" or other			EnvVariable:	LSlvMark 				: ">" is default
	'		2c. Warninf if TClogFile not open EnvVariable:	TClogWarning	-w, -W *: True, False (default)
	'			Note: Explicitly given -s or -l imply -w, and then explicit -W is required to disable warnings
	' Other options: -d - apply current options as global defaults (can be given without a cmdline proper)
	' Error exit if contradictory options entered
	' Note also: possibility to go into "direct-slave mode", ie.
	' all input and output goes directly to slave and back, plus an
	' escape command to bring us back to normal

	Dim opts, opt_echo, opt_log
	opts = ""
	saydbgq "@ssend calling getopts with :El, cmdline=" & cmdline
	getopts ":El", cmdline, opts
	saydbgq "@ssend after calling getopts, opts=" & opts

	opt_echo	=  True: If find_opt("E", opts) then opt_echo	= False
	opt_log		= False: If find_opt("l", opts) then opt_log	= True
	
	if SlaveExec Is Nothing then
		sayerr "Error: No slave app to send to."
		Exit Sub
	End If
	if opt_log Then TClog "(ssend) Sending: " & cmdline
	
	saydbg "@ssend Sending: " & cmdline
	SlaveExec.StdIn.WriteLine(cmdline)
	FoundLine = SlveReadUpto (SlavePrompt)
	if opt_echo And len(FoundLine) > 0 Then
		oRx.global = true
		' oRx.ignorecase = false		' not needed
		oRx.pattern = VBCrLf&"|\n"
			TClog SlaveOutFlag & oRx.Replace (FoundLine, VBCrLf & SlaveOutFlag)
'			say SlaveOutFlag & oRx.Replace (FoundLine, VBCrLf & SlaveOutFlag)
	End If

End Sub ' ssend(cmdline)
' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O


' What about: SayLocVars ("-L -d myfunction -q"" v1 v2 v3")
' 	- option -d myfunction creates: saydbg "@myfunction <rest of line>"
'	- option -d without value creates: saydbg "..."
'	- option -d missing creates: say "..."
' 	- option -r writes to stderr instead of stdout, but is overridden by -d
' 	- option -L writes everything on the same line, normal is one per line
'	- option -q<ch> "quotes" the output value using character <ch> as quotes
'				Only first char of <ch> is used, rest is ignored, or
'	 			eg. "-q/ var" would produce: var=/<value of var>/


Set vSayLvar = Getref ("vbsCodetoSayLocalVar") ' short form, Usage: execute vSayLvar ("myvar")

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
Function vbsCodetoSayLocalVar (v)
	' Helps with debugging and saves you some tricky typing like: say "myvar="&myvar
	' Usage: Use the following syntax: execute vbsCodetoSayLocalVar("myvariable")
	' 		to print out "myvariable=<whatever is in myvariable>"
	' NOTE: This function does not print anything, it builds the code you can "execute" to print LOCAL
	'		variables that are ONLY visible in the environment you are calling from.
	'		It works for global vars too, unless there is a local var with the same name.
	vbsCodetoSayLocalVar = "say """ & v & "="" & eval(""" & v & """)"
	'sayerr "vbsCodetoSayLocalVar ='"&vbsCodetoSayLocalVar&"'"
End Function ' Function vbsCodetoSayLocalVar (v)

Set vSayLvarq = Getref ("vbsCodetoSayLocalVarq") ' short form, Usage: execute vSayLvarq ("myvar")

' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
Function vbsCodetoSayLocalVarq (v)
	' Helps with debugging and saves you some tricky typing like: say "myvar='"&myvar&"'"
	' Usage: Use the following syntax: execute vbsCodetoSayLocalVarq("myvariable")
	' 		to print out "myvariable='<whatever is in myvariable>'"
	' NOTE: Same as vbsCodetoSayLocalVar above, but the '<value of var>' is enclosed in single quotes
	vbsCodetoSayLocalVarq = "say """ & v & "='"" & eval(""" & v & """)" & "&""'"""
	sayerr "vbsCodetoSayLocalVarq ='"&vbsCodetoSayLocalVarq&"'"
End Function



'_h2 #DBG_ routines - Selective debug messages: enabled/disabled by topic or function

_ 
' Globals related to saydbg(s)
' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
	Dim DBG_enabled, DBG_current, DBG_banner, DBG_bannerDefault, DBG_RecurSTOP
	DBG_bannerDefault = "#### DEBUG"
	If vartype (DBG_enabled) = 0 Then	' Only init first time: Conserve vars if this code is reimported
		DBG_enabled		= "|"		' holds enabled debug topics separated by vertical bar eg. "|topic1|topic2|...""
		DBG_topic		= "|"
		DBG_banner		= DBG_bannerDefault
		DBG_RecurSTOP	= False		' Switch on temporarily to prevent infinite recursion
	End If

 ' dbg (dbgCmd) - Manage debug settings
 ' ====================================================================================================
 Sub dbg (dbgCmd)
	Dim cmd
	if dbgCmd = "nonexistantrubbish dsafdsfdafadsfklj" Then
		cmd = "help"
	ElseIf False Then ' DEAD CODE
		dbgCmd = GcmdLine
		sayerr "Setting dbgCmd - GcmdLine='"&GcmdLine&"'"
		cmd = readToken(dbgCmd)
		cmd = readToken(dbgCmd)
		sayerr "cmd='"&cmd&"'"

	End If

	cmd = readToken (dbgCmd)
	if cmd = "" Then cmd="?"
	if sDBcheck ("|help|h|?|", cmd) Then
		say "  dbg - Manage debug messages from the saydbg routine, based on enabling/disabling keywords"
		say "  The command: saydbg ""@keyword <message>"" - Will output <message>only if keyword is enabled"
		say "  Usage: .dbg <args> ... where args are:"
		say "       h[elp]|?               - Show this help text"
		say "       show                   - Show globals: DBG_enabled, DBG_topic, DBG_banner"
		say "       add <keywords> ...     - Add/enable one or more keywords (tokens)"
		say "       add all                - enable ALL debug messages"
		say "       del[ete] <keywords>... - Delete/disable one or more keywords"
		say "       reset                  - Reset ie. disable all"
		

	Elseif sDBcheck ("|show|", cmd) Then
		' show
		If DBG_RecurStop Then
			Say "dbgsay is DISABLED"
		Else
			Say "dbgsay is ENABLED"
		End If
		sayvarq "DBG_enabled"
		sayvarq "DBG_topic"
		sayvarq "DBG_banner"
		

	Elseif sDBcheck ("|reset|", cmd) Then
		' reset
		say "dbg reset:"
		DBG_enabled = "|"
		DBG_topic   = "|"
		DBG_banner  = DBG_bannerDefault
		DBG_RecurSTOP  = False

	Elseif sDBcheck ("|add|", cmd) Then
		t="  "
		do while Not t = ""
			t = readToken (dbgCmd)
			if left(trim(t),1) = "'" then exit do
			sDBadd DBG_enabled, t
		Loop

	Elseif sDBcheck ("|delete|del|", cmd) Then
		t="  "
		do while Not t = ""
			t = readToken (dbgCmd)
			sDBdelete DBG_enabled, t
		Loop
	
	Else 
		' not recognized. Give the help message by default
		dbg "help"
	End If
 End Sub ' Sub dbg (dbg_cmd)

Dim oRxDbgTopic	: Set oRxDbgTopic = NewRegExp("^[^\S\n]*\@?(\.?\S*)[^\S\n]?([\s\S]*)$", True, False)

Function dbg_ontopic (ByRef s)
 	Dim tok
	dbg_ontopic = ""		' By default

	If sDBcheck (DBG_enabled, "all") Then		' "all" globally enabled means we are always on-topic
		dbg_ontopic = "all"
	Else
		tok = oRxDbgTopic.Replace (s, "$1")
		For Each t in Split(tok, "|")
			If sDBcheck (DBG_enabled, t) Then
				If dbg_ontopic <> "" Then dbg_ontopic = dbg_ontopic + ","
				dbg_ontopic = dbg_ontopic + t
			End If
		Next
	End If

End Function

Sub saydbgq (ByRef s)
	saydbg ("-q " & s)
End Sub

 Sub saydbg (ByRef s)
	' TODO: Selective debug messages: enabled/disabled by topic or function
	' Global vars:  DBG_enabled = /abc/def/regex/...  these ones are enabled
	'				DBG_current = /regex/ - current topic, ie, must be found
	'				alternatively, the topic can be passed in as parameter
	'				signalled by initial at-sign dash eg. saydbg "@regex <message> "
	Dim topic, topics, t, on_topic, opts, opt_quoted,s2
	
 	If DBG_RecurSTOP Then Exit Sub		 ' Recursive call is stopped here
	DBG_RecurSTOP  = True				 ' Temporarily disable this Sub, to prevent infinite recursion (eg. from getopts)
										 ' ie. Any routines called from here (getopts in particular) are temporarily
										 ' shut out and denied the use of this functionality. All calls will be silently ignored.
										 ' NOTE: Debug messages from getopts will still work perfectly OK in most cases
										 '		 Only the calls to getopts made from within here have debugging disabled.

	topic = ""
	getopts ":q", s, opts
	opt_quoted = False
	if find_opt("q", opts) then opt_quoted = True

	DBG_disable  = False

	' Get topic from s, if present, in the form "@topic <rest of string>"
	If InStr (rxLTrim(s),"@") = 1 Then
		sv    = split(rxLTrim(s), " ", 2)
		topic = Trim(Replace(sv(0),"@", "", 1, 1))
		if ubound(sv) >= 1 Then
			s = sv(1)
		Else
			s = ""
		End If
	End If
	topics = "|" & topic & "|" & DBG_topic 
	' TODO: Decide: Should empty s be rejected here ? Currently it isn't
	
	' Now check if we are on-topic
	If sDBcheck (DBG_enabled, "all") Then		' "all" globally enabled means we are always on-topic
		'WScript.Stderr.WriteLine DBG_banner & "(saydbg): ALL enabled: Forced on_topic"
		on_topic = True
	Else
		on_topic = False ' By default
		For Each t in Split(topics, "|")
			t = Trim (t) ' Probably not necessary but seems right
			If sDBcheck (DBG_enabled, t) Then
				'WScript.Stderr.WriteLine DBG_banner & "(saydbg): on_topic enabled, key: " & t
				on_topic = True
				Exit For
			End If
		Next
	End If

	' If Not on_topic Then WScript.StdErr.WriteLine DBG_banner & "(saydbg)  NOT on topic: " & s
	
	If on_topic Then
		bann = DBG_banner & ": "
		' If len (topic) > 0 Then bann = DBG_banner & "(" & topic & "): "
		If len (t) > 0 Then bann = DBG_banner & "(" & t & "): "
		
		If opt_quoted Then s= "'" & s & "'"
		
		WScript.StdErr.WriteLine bann & s
	End If
	DBG_RecurSTOP  = False		' Re-enable this Sub ie. Lift the recursion protection, as we are done now
 end sub ' Sub saydbg (s)
'# End DBG Routines
_ 

' help () - Main Help Routine
' ====================================================================================================
Sub help ()
	Dim nl: nl=VBCrLf
	say "" _
	&"    h[elp]|?              - This help text" & nl _
	&"    q|quit                - Quit" & nl _
	&"    say ""text ...""        - Print text to console" & nl _
	&"   .say text ...          - same as above, but text is preprocessed and enclosed in quotes" & nl _
	&"    cmd text ...          - Run a VBScript command line interactively" & nl _
	&"   .cmd text ...          - same as above, but text is preprocessed and enclosed in quotes"

End Sub ' Sub help ()


' readToken (ByRef s) - Get first token from a command line
' ====================================================================================================
Function readToken (ByRef s)
	Dim arr
	readToken = ""
	If RxTrim(s) = "" Then Exit Function		' empty s, no tokens there, just quit

	arr 		= Split (s, " ", 2)				' split into <first token>, <rest of line>
	readToken	= ""
	s 			= ""
	readToken 	= arr (0)						' arr has at least one element
	If ubound (arr) > 0 Then s = arr (1)		' If two, put second one in s

End Function ' Function readToken (s)



' ====================================================================================================
'_h1 Misc funtions
' ====================================================================================================

' dotargs - Run a command line, after substitution  (_maybe_ useful in debugging)
' ====================================================================================================
' TODO: Not clear how this is useful. Consider removing
Sub dotargs (s)
	RunVbshLine preprocess_cmdline (s)
End Sub

' sayvar - Echo (global) variable name and value (useful in debugging)
' ====================================================================================================
' TODO: The _global_ requirement is pretty restrictive.
'		This would be much better if implemented as a preprocessor macro
Sub sayvar(v)		' NOTE: v must be the name of a GLOBAL variable
	say v & "="  & eval(v)
End Sub

' sayvarq - Echo (global) variable name and value within single quotes (useful in degugging)
' ====================================================================================================
Sub sayvarq (v)		' NOTE: v must be the name of a GLOBAL variable
	say v & "=" & "'" & eval(v) & "'"
End Sub

' Check if Blawber.exe has been started as slave app
' ====================================================================================================
' TODO: Not sure if this is the right way. Better implement something general
Function Blawber_started()
	If (SlaveFname = "") or (SlaveShell = Nothing) or (SlaveExec = Nothing) then 
		Blawber_started = False
	Else
		Blawber_started = True
	End If
End Function ' Blawber_started()


' ====================================================================================================
'_h1 sDB--- "silly DataBase" routines (sDBadd, sDBdelete, sdBcheck, sDBreset)
' Store keywords in a string variable, separated by | , similar to a PATH variable
' ====================================================================================================
_
 Function sDBcheck (ByRef sDBstring, key)
 	sDBcheck = False
 	If len (key) = 0 Then Exit Function
 	If InStr (sDBstring, "|" & key & "|") > 0 Then sDBcheck = True
 End Function ' Function sDBcheck (ByRef sDBstring, key)
 
 Sub sDBdelete (ByRef sDBstring, key)
 	If len (key) = 0 Then Exit Sub
 	sDBstring = Replace(sDBstring,"|" & key & "|","|") 
 End Sub ' Sub sDBdelete (ByRef sDBstring, key)
 
 Sub sDBadd (ByRef sDBstring, key)
 	If len (key) = 0 Then Exit Sub
 	sDBstring = Replace(sDBstring,"|" & key & "|","|") 'Delete first, to avoid doubles
 	sDBstring = sDBstring & key & "|"
 End Sub ' Sub sDBadd (ByRef sDBstring, key)
 
 Sub sDBreset (ByRef sDBstring)
 	sDBstring = "|"
 End Sub ' Sub sDBreset (ByRef sDBstring)
_


' tgopts (cmdline) - Test getopts (use in debugging)
' =================================================================================================
Sub tgopts (cmdline)
	saydbgq "@tgopts initial optstr=" & optstr
 	If optstr = "" Then optstr = "abcde:f:x:z|good-times|parse-slowly:|a-good-opt:|"
	saydbgq "@tgopts used optstr=" & optstr
	opts_found = ""

	getopts optstr, cmdline, opts_found


	saydbg "@tgopts Result: ========================================="
	saydbg "@tgopts opts_found  :" & opts_found
	saydbg "@tgopts Rest of line:" & cmdline
	g_opts_found = opts_found
	saydbg "@tgopts g_opts_found  :" & g_opts_found

	' Never mind the clumsy variables, use find_opt and find_opt_val instead
	Exit Sub

	' DEAD Code Below this point. Left in for reference in case needed later
	' =================================================================================================

	saydbg "@tgopts calling split_optarr"
	arrOpts = split_optarr(opts_found)
	for i = lbound(arrOpts) to ubound (arrOpts)
		saydbg "arrOpts("&i&")="&arrOpts(i)
		execute arrOpts(i)
	next

	saydbg "@tgopts *****************************  Checking variables set ************"

	for i = lbound(arrOpts) to ubound (arrOpts)
		saydbg "@tgopts arrOpts("&i&")="&arrOpts(i)
		myname = left(arrOpts(i), instr(arrOpts(i),"=")-1)
		'myline = """say arrOpts("""&i&""")  \&"""myname
		myline = "say " & myname
		saydbg "@tgopts running: "&myline
		execute myline
	next

End Sub ' Sub tgopts (cmdline)

'_h2 find_opt(opt, found_opts) - Checks for presence of opt and returns true or false accordingly
' =================================================================================================
Function find_opt(opt, found_opts)
	if Instr(opt,"-") = 1 Then opt = Mid(opt,2) ' remove one or two leading dashes
	if Instr(opt,"-") = 1 Then opt = Mid(opt,2)
	find_opt = false ' by default
	If   InStr(found_opts, "|"&opt&"|") > 0 _
	Or   InStr(found_opts, "|"&opt&"=") > 0 _
	Then find_opt = true
End Function

'_h2 find_opt_val(opt, found_opts) - Checks for presence of opt and returns its arg(value) if any
' =================================================================================================
Function find_opt_val(opt, found_opts)
	if Instr(opt,"-") = 1 Then opt = Mid(opt,2) ' remove one or two leading dashes
	if Instr(opt,"-") = 1 Then opt = Mid(opt,2)
	find_opt_val = "" 							' by default
	If InStrRev(found_opts, "|"&opt&"|") > 0 Then Exit Function
	
	pos = InStrRev(found_opts, "|"&opt&"=")
	if pos = 0 Then Exit Function
	
	prev = InStr(pos, found_opts, "=") + 1
	pos  = InStr(prev, found_opts, "|")

	val = Mid(found_opts,prev, pos-prev)
	find_opt_val=val
End Function

'_h2 opt_sgClash(opts, found_opts) - Checks one or more single-letter opts and returns the ones found present
'	eg. If opt_sgClash("Ffg", found_opts) returns "Fg" then both -F and -g were given (which may be conflicting options)
' =================================================================================================
Function opt_sgClash (opts, ByRef found_opts)
	Dim optchar
	opt_sgClash = "" ' by default

	For i = 1 To len (opts)
		optchar = Mid(opts, i, 1)
		If   InStr(found_opts, "|"&optchar&"|") > 0 _
		Or   InStr(found_opts, "|"&optchar&"=") > 0 _
		Then opt_sgClash = opt_sgClash & optchar
	Next
End Function

'_h2 opt_sgClashCsv(opts, found_opts) - Calls opt_sgClash for each element in a csv-list and quits on first hit
'	Where a "hit" means conflicting options found, eg. "Fgh" results in a value beginning with "F" and longer than 1,
'	A return value of length 0 or 1 is never a "hit"
'	A return value of any length is only a hit if the first char is same as in the input string (eg. "F" in the example)
' =================================================================================================
Function opt_sgClashCsv (opts_csv, ByRef found_opts)
	Dim opts_arr, opts, optsCh1, sgClash
	opt_sgClashCsv = "" ' by default
	opts_arr = Split (opts_csv, ",")


	For i = LBound(opts_arr) To UBound(opts_arr)
		opts = Trim(opts_arr(i))
		optsCh1 = Mid(opts, 1,1)
		sgClash = opt_sgClash (opts, found_opts)

		If  Len (sgClash) > 1 _
		And Mid (sgClash,1,1) = Mid(opts, 1,1) Then 
			opt_sgClashCsv = sgClash
			Exit Function
		End If
	Next
End Function

'_h2 function split single-letter_optarr (s) - After getopt, split options into an array
'	eg. If function spl"Ffg"ingle-letter_ returns "Fg" then both -F and -g were given (which may be conflicting options)
' =================================================================================================
' TODO: test function split_optarr and see whether it can be useful or not
'		currently not used
function split_optarr (s)
	'Dim split_opts()
	split_opts = split (s,"|")
	saydbg "ubound(split_opts)="&ubound(split_opts)
	for i = lbound(split_opts) to ubound (split_opts)
		str = Replace (split_opts(i), "-", "_")
		eqpos = InStr(str, "=")
		if eqpos <= 0 then
			split_opts(i) = "opt_" & str & "=""Y"""
		else
			name = Left(str, eqpos - 1)
			val = Mid (str, eqpos + 1)
			if len(name) > 1 then name = "_" & name
			if IsNumeric (val) then
				split_opts(i) = "opt_" & name & "=" & val
			else
				split_opts(i) = "opt_" & name & "=""" & val & """"
			end if
		end if
	next
	split_optarr = split_opts
end function

'  =================================================================================================
' Get options from a command line
' =============================================================================
'! Parse Unix-style command line options
'! @param optstr     - String of valid options eg. ":abc"
'! @param cmdline    - The command line to parse eg. "cmd -abf file arg1 arg2 ..."
'! @param opts_found - Options found and their values are returned in this string
Sub getopts (optstr, ByRef cmdline, ByRef opts_found)
	' optstr 	- string of accepted options eg. "abc:dD:st" where
	'			  c and D require an argument such as a filename 
	'			  NOTE: --long-style-opts can be included at end of optstr
	'					with double-dash removed and separated by vertical bar
	'					eg."abc:|long-style-opt1|long-style-opt2:|..."
	' cmdline	- the line to be parsed eg. "cmd -bc testfile.tst infile outfile"
	'			  after parsing the cmd and opts found are stripped away and
	'			  cmdline contains just the args eg. "cmd testfile.tst ..."
	' opts_found - string where detected opts and their args are stored
	' =============================================================================
	
	' Output, options result string:
	' opts_found: ":opt_a:opt_b=abcde:opt_c:#opt_x#opt_y# ..."
	' Where expected opts (as per optstr) are enclosed in colon :opt_expected:
	' and unexpected opts (not present in optstr) are enclosed in hash mark #opt_unexpected#
	Dim oRegOpts, oMatch, cmd_token
	cmd_token = ""

	' Pattern for short and long opts including optvalues
	sPat = "^\s*((-[a-zA-Z]+(=""[^""]*""|='[^']*'|=\S*)?)|(--\w[-\w]*(=""[^""]*""|='[^']*'|=\S*)?))(.*)$"
	sPat = "^\s*((-\w+(=""[^""]*""|='[^']*'|=\S*)?)|(--\w[-\w]*(=""[^""]*""|='[^']*'|=\S*)?))(.*)$"

	'Pattern for optvalue. Currently not used
	vPat = "^\s*(((""[^""]*""|='[^']*'|\S*)?))(.*)$"

	restline = cmdline
	sep_char = "|"			' separator character used in opts_found eg. "|a|b|z|f|"
	opts_found = sep_char
	long_found = sep_char

	' 	optstr = "abcde:x:z|good-times|parse-slowly:|a-good-opt:|"
	sep_charpos = InStr(optstr, sep_char)
	if sep_charpos > 0 Then
		short_optstr = Left (optstr, sep_charpos - 1) & "  "	' add two spaces as out-of-range protection
		long_optstr  = Mid  (optstr, sep_charpos    ) & "  "	' add two spaces as out-of-range protection
	Else
		short_optstr = optstr & "  "	' add two spaces as out-of-range protection
		long_optstr  = "  "	' add two spaces as out-of-range protection
	End If

	std_errors = true			' Error exit: opt unknown, opt value missing when required or given when not
	strict_errors = false		' Error exit: opt repeated

	If InStr(short_optstr,"!") = 1 Then strict_errors = true
	If InStr(short_optstr,":") = 1 Then std_errors = false

	Set oRegOpts = New RegExp
	oRegOpts.global = true
	oRegOpts.ignorecase = false
	
	saydbg "@getopts Initial restline='" & restline & "'"

	'oRegOpts.pattern = "^\s*([a-zA-Z]+\w*)?"				' Regex to remove initial command token, if any
	'oRegOpts.pattern = "^(\s*[^-\s][\S]*[\s]*)?"				' Regex to remove initial command token, if any
	'oRegOpts.pattern = "^\s*([^-\s][\S]*)?\s*"				' Regex to remove initial command token, if any
	oRegOpts.pattern = "^(\s*[^-\s][\S]*)?"				' Regex to remove initial command token, if any
	' NOTE: cmd token = wspace plus ANY non-blank sequence that does NOT begin with dash plus trailing wspace
	' anything up to first <-option> or second token whichever comes first

	Set oMatch = oRegOpts.Execute(restline)
	If oMatch.Count > 0 Then
		cmd_token = oMatch(0).submatches(0)
	End If
	saydbg "@getopts cmd_token='" & cmd_token & "'"
	
	restline = oRegOpts.Replace (restline, "") ' remove cmd token
	 ' TCase: What if cmd token is repeated somewhere in restline. Ensure only the first one is removed


	saydbg "@getopts cmd removed, restline='" & restline & "'"

	oRegOpts.pattern = sPat

	Do
		saydbg "@getopts_do1 opts_found:" & opts_found			' **** DEBUG
		saydbg "@getopts_do1 Checking:" & oRegOpts.pattern			' **** DEBUG
		saydbg "@getopts_do1 Against :" & restline			' **** DEBUG
 
		Set oMatch = oRegOpts.Execute(restline)
		
		'saydbg "@getopts_do2 oMatch.Count=" & oMatch.Count & " oMatch.submatches.Count=" & oMatch.submatches.Count
		'saydbg "@getopts_do2 oMatch.Count=" & oMatch.Count
		If oMatch.Count <= 0 Then Exit Do

		'saydbg "@getopts_do2 oMatch(0).submatches.Count=" & oMatch(0).submatches.Count

		'For j = 0 to oMatch(0).submatches.Count - 1			' **** DEBUG
		'	say "		oMatch(0).submatches("&j&")=" & oMatch(0).submatches(j)			' **** DEBUG
		'Next			' **** DEBUG

		If oMatch(0).submatches.Count < 2 Then Exit Do

		max = oMatch(0).submatches.Count-1
		restline = oMatch(0).submatches(max)
		saydbg "@getopts Assigned restline='"&restline&"'"
		this_tok = oMatch(0).submatches(0)
		this_opt=""

		If InStr (this_tok, "--") = 1 Then		' handle long opts eg. --verify-inputs-type=strict
			this_tok =  mid(this_tok, 3)

			this_optv=""
			this_optvpos = InStr (this_tok, "=")

			If this_optvpos  > 0 Then
				this_optv =  mid(this_tok,this_optvpos)
				this_tok  = left(this_tok,this_optvpos-1)
				this_optv =  Replace(this_optv,"""", "")
				this_optv =  Replace(this_optv,"'", "")
				'sayerr "		this_optvpos=" & this_optvpos			' **** DEBUG
				'sayerr "		this_optv   =" & this_optv				' **** DEBUG
				'sayerr "		this_tok    =" & this_tok				' **** DEBUG
			End If

			this_opt = this_tok

			if (strict_errors) and ( _
					(InStr(opts_found, sep_char & this_opt & sep_char) > 0 ) _
					or 	(InStr(opts_found, sep_char & this_opt & "=") > 0 ) _
			 	) Then
				sayerr "Error: repeated option -" & this_opt
				Exit Do
			end if

			If InStr(long_optstr,sep_char & this_opt&":") > 0 Then
				' opt recognized, opt value required
				If len (this_optv) > 0 Then
					' Value given as --this-opt=value
					opts_found = opts_found & this_opt & this_optv & sep_char
				ElseIf std_errors Then
					sayerr "Error: option requires value --" & this_opt
					Exit Do
				Else
					opts_found = opts_found & this_opt & sep_char
				End If

			ElseIf InStr(long_optstr,sep_char & this_opt & sep_char) > 0 Then
				' opt recognized, opt value not required
				If len (this_optv) > 0 Then
					If std_errors Then
						sayerr "Error: option does not require value --" & this_opt
						Exit Do
					Else
						' With errors disabled, accept the value anyway
						opts_found = opts_found & this_opt & this_optv & sep_char
					End If
				Else
					opts_found = opts_found & this_opt & sep_char
				End If

			Else
				' opt unknown ie. not present in short_optstr
				If std_errors Then
					sayerr "Error: invalid option --" & this_opt
					Exit Do
				end if
				' opt unknown is accepted with errors disabled
				opts_found = opts_found & this_opt & sep_char


			End If

		Else									' handle short opts eg. -abc=outfile.log

			If Instr (this_tok, "-") = 1 Then this_tok = mid(this_tok,2)

			this_optv=""
			this_optvpos = InStr (this_tok, "=")

			If this_optvpos  > 0 Then
				this_optv =  mid(this_tok,this_optvpos-1)
				this_tok  = left(this_tok,this_optvpos-2)
				this_optv =  Replace(this_optv,"""", "")
				this_optv =  Replace(this_optv,"'", "")
				'sayerr "		this_optvpos=" & this_optvpos			' **** DEBUG
				'sayerr "		this_optv   =" & this_optv				' **** DEBUG
				'sayerr "		this_tok    =" & this_tok				' **** DEBUG
			End If

			for i = 1 to len (this_tok)
				this_opt=mid(this_tok,i,1)
				this_pos = InStr(short_optstr,this_opt)

				'sayerr "**** DEBUG: handling opt -" & this_opt
				'sayerr "**** DEBUG: short_optstr=" & short_optstr


				if (strict_errors) and (InStr(opts_found, sep_char & this_opt) > 0 ) Then
					sayerr "Error: repeated option -" & this_opt
					Exit Do
				end if

				if InStr(short_optstr,this_opt&":") > 0 Then
					' opt recognized, opt value required
					'saydbg "@getopts opt recognized, opt value required -" & this_opt		' ****DEBUG:
					If len (this_optv) > 0 Then
						' -abcdf=value where f is last in token, and value given as "=value"
						opts_found = opts_found & this_optv & sep_char
					ElseIf i = len (this_tok) then
						' -abcdf=value where f requires value, is last in this token
						this_opt = "-" &this_opt & "=" 			' disguise "f" as "-f="
						restline = this_opt & rxLTrim(restline)		' stick it in before rest of line
						saydbg "@getopts Assigned Disguised restline:'"&restline&"'"		' **** DEBUG

						Exit For								' and simply handle in next iteration
					Else	
						' -abfcd where -f requires value but is not last in this token
						if std_errors then
							' -abfcd is invalid if -f requires value
							sayerr "Error: option requires value -" & this_opt
							Exit Do
						Else
							' however, with errors disabled, just accept f without a value
							opts_found = opts_found & this_opt & sep_char
						end if
					End If

				Elseif InStr(short_optstr,this_opt) > 0 then
					' opt recognized, opt value not required
					If (i = len (this_tok)) and (len (this_optv) > 0) then

						if std_errors then
							sayerr "Error: option does not require value -" & this_opt
							Exit Do
						Else
							' however, with errors disabled, just accept f with a value
							opts_found = opts_found & this_opt & this_optv & sep_char
						end if
					Else
						opts_found = opts_found & this_opt & sep_char
					End If
				else
					' opt unknown ie. not present in short_optstr
					If std_errors Then
						sayerr "Error: invalid option -" & this_opt
						Exit Do
					end if
					' opt unknown is accepted with errors disabled
					opts_found = opts_found & this_opt & sep_char

				End If

			next
			if len (this_optv) > 0 Then opts_found = opts_found & this_optv & sep_char
		End If
	
	loop while oMatch.Count > 0
	
	'Remove one initial space from restline if present
	If Instr (restline, " ") = 1 Then restline = Replace (restline, " ", "", 1, 1)
	
	saydbg "@getopts Result: ========================================="
	saydbg "@getopts opts_found  :" & opts_found
	saydbg "@getopts Rest of line:" & restline
	
	If len(opts_found) < 2 Then 
		' Do nothing: No opts found, so DON'T CHANGE cmdline AT ALL
	ElseIf len(cmd_token) > 0 Then 
		cmdline = cmd_token & " " & restline
	Else
		cmdline = restline
	End If
	saydbg "@getopts resulting cmdline:'" & cmdline & "'"

End Sub ' Sub getopts (optstr, ByRef cmdline, ByRef opts_found)

' getIniOpts - Get _Initial_ options: As getopts, but options must come first.
' initial (command) token in NOT recognized or handled in any way
Sub getIniOpts (optstr, ByRef line, ByRef opts_found)
	opts_found = ""
	If Left(RxLTrim(line),1) = "-" Then getopts optstr, line, opts_found
End Sub



set argsubst0 = Getref("replace_args")
sub sppa(s): say replace_args(s): end sub		' *** DEBUG

' ====================================================================================================
Function replace_args (argline)
	' Function replace_args (argline)
	' Preprocess args from a command line of format "<command> <rest of line>":
	' 	1. The first token <command> should have been removed,
	'	   only the part <rest of line> is processed here
	' 	2. {expr} found anywhere in <rest of line>, except within quotes "",
	'	   is substituted for the return value of the vbscript expression expr
	'	   expr can be a simple variable name or a more complicated expression
	'	   eg. {"Date: " & date & " Time: " & time}
	'	3. {expr ... gives error message and exits if closing curly brace missing
	'	4. Anything within "double quotes" is preserved, ie. subsitution of
	'	   {expr} does not happen, and missing closing brace is not detected
	'	5. double quotes ""are doubled"", producing normal (undoubled) quotes
	'	   when the resulting line is further processed as a string
	'	6. Finally, "<rest of line>" is returned, enclosed in quotes
	'	NOTE: The resulting line is NOT right-trimmed, (and not even left anymore, see below)
	'	NOTE: Left-trimming was also removed, but not properly tested yet: 2018-06-15 V0.05-01
	'	TODO: Test whether removal of Left-trimming affected anything
	
	'Declare variables to make sure they are local
	'Note EXCEPT GtempVar that is assigned by ExecuteGlobal and NEEDS to be global
	Dim f_name, f_error, remline, argl2, pq, ps, s, cmline

	
	
	replace_args = argline			' Inputline is returned as is in case of error exit
	f_name = "replace_args"
	f_error = "Error(" & f_name & "): "
	' say "replace_args called, argline="&argline		' *** DEBUG
	
	' argline  = trim(argline)
	if len(Trim(argline)) = 0 then Exit Function		' Empty line is OK, we just quit silently
	
	'remline = rxLTrim(argline)
	remline = argline
	argl2 = ""

	do while len  (remline) > 0
		pq = InStr(remline,"""")  ' position of first quote character
		ps = InStr(remline,"{")  ' position of first "{" denoting substitution 
	
		if (ps > 0 and (pq = 0 or pq > ps)) then		' substitution found first
			argl2 = argl2 & left (remline, ps - 1)
			remline = mid (remline, ps+1)
			ps = InStr(remline,"}")
			if ps = 0 then
				sayerr f_error & "Missing right curly brace ""}"" in cmdline: " & cmdline
				Exit Function
			end if 
			s = left (remline, ps - 1)
			remline = mid (remline, ps + 1)
			' say "remline=/"&remline&"/"											' *** DEBUG
	
			On Error Resume Next
			Err.Clear
			'say "ExecuteGlobal (""GtempVar=""&trim("&s&"))"		' *** DEBUG
			'ExecuteGlobal "GtempVar="&trim(s)
			ExecuteGlobal "GtempVar="""" & "&trim(s)
			' GtempVar = Eval(s)
			If Err.Number <> 0 Then
				sayerr f_error & "Unable to substitute variable: """ & trim(s) & """"
				sayerr Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
				GtempVar=""
				If strictErrExit Then		' strictErrExit means one error causes the whole line to be left unparsed
					Exit Function
				Else 
					' GtempVar="{" & s & "}" 	' putting the erroneous string back in was an experiment, but not a good idea.
					' This way {<anything unvalid>} is just removed, including the curlies
				End If
			End If
			On Error Goto 0

			argl2 = argl2 & GtempVar
		elseif (pq > 0 and (ps = 0 or ps > pq)) then  ' quote found first
			argl2 = argl2 & left (remline, pq) & """"
			remline = mid (remline, pq + 1)
			pq = InStr(remline,"""")
			if pq = 0 then
				s = remline & """"""		' missing end quote is ok, we just supply it
				remline = ""
			else
				s = left(remline, pq) & """"
				remline = mid (remline, pq + 1)
			end if
			argl2 = argl2 & s
		else
			argl2 = argl2 & remline
			remline = ""
		end if	
		' say "remline=/"&remline&"/"		' *** DEBUG
	
	loop
	'say "cmd="&cmd		' *** DEBUG
	'say "argl2="&argl2		' *** DEBUG
	cmline = argl2

	' say "Resulting command line=/"&cmline&"/"		' *** DEBUG

	replace_args = argl2

	' ====================================================================================================
End Function ' Function replace_args (argline)

' Position of first non-space char in s
' or = len(s) + 1 if none
Function Ltrimpos (s)
	LTrimPos = len (s) - len (rxLtrim(s)) + 1
End Function

' ====================================================================================================
Function preprocess_cmdline (cmdline)
	' Function preprocess_cmdline (cmdline)
	' Preprocess a command line of format "[.]<command> <rest of line>" as follows:
	' 	0. An initial "." dot is removed if found, along with any spaces around it.
	' 	1. The first token <command> is preserved and the <rest of line> is processed
	' 	2. {expr} found anywhere in <rest of line>, except within quotes "",
	'	   is substituted for the return value of the vbscript expression expr
	'	   expr can be a simple variable name or a more complicated expression
	'	   eg. {"Date: " & date & " Time: " & time}
	'	3. {expr ... gives error message and exits if closing brace missing
	'	4. Anything within "double quotes" is preserved, ie. subsitution of
	'	   {expr} does not happen, and missing right brace is not detected
	'	5. double quotes ""are doubled"", producing normal (undoubled) quotes
	'	   in the further processing
	'	6. Finally, <rest of line> is enclosed in quotes and
	'	   <command> "<rest of line>" is returned, making it suitable
	'	   for being run directly, if <command> was defined as a Sub
	'	   TODO: Adapt to work with multiple arguments
	'	   TODO: Adapt this, or write separate routine to call functions
	'			ie. where arguments need to be enclosed in brackets
	preprocess_cmdline = ""
	f_name = "preprocess_cmdline"
	f_error = "Error(" & f_name & "): "
	' say "preprocess_cmdline called, cmdline="&cmdline		' *** DEBUG

	cmd = ""
	argline  = ""

	' Remove initial "." dot if present, but not spaces before or after dot
	' cmdline = rxLTrim (cmdline)
	ltpos = LTrimPos (cmdline) ' position of first non-blank
	'if Left (rxLTrim (cmdline), 1) = "." Then cmdline = rxLTrim (Mid (cmdline, 2))


	if Mid (cmdline, ltpos, 1) = "." Then cmdline = Replace (cmdline, ".", "", 1, 1)
	ltPos = LTrimPos (cmdline) ' new position of first non-blank, after dot removed
	If ltPos > len (cmdline) Then 		' Blank or empty cmd is OK, just quit silently
		preprocess_cmdline = cmdline
		Exit Function
	End If

	'sayerr "preprocess_cmdline Got here 2"				' *** DEBUG

	remline = mid (cmdline, ltPos)

	cmd = Left (cmdline, ltpos-1) & readToken (remline)
	argline = remline
	
	'say "cmd=/"&cmd&"/"				' *** DEBUG
	'say "argline=/"&argline&"/"		' *** DEBUG
	
	argl2 = replace_args (argline)

	'say "cmd="&cmd		' *** DEBUG
	'say "argl2="&argl2		' *** DEBUG
	cmline = cmd & " " & """"&argl2&""""

	saydbg "@preprocess_cmdline Resulting command line=/"&cmline&"/"		' *** DEBUG

	preprocess_cmdline = cmd & " " & """"&argl2&""""

End Function ' Function preprocess_cmdline (cmdline)


'! Find the first occurrence of the given filename from the working directory
'! or any directory in the %PATH%.
'!
'! @param  filename   Name of the file to find.
'!
Private Function findFileName(ByVal filename)
	Dim fso, sh, file, code, dir, opts_found, opt_n, fxtensions
	' opt -n : accept non-existant filename as if it existed in current dir
	'			otherwise return empty string
	' opt -x .ext1.ext2....	: Filename extensions to try if filename lacks one and is not found
	'
	opts_found = "" 
	getopts ":nx:", filename, opts_found
	opt_n = find_opt ("n", opts_found)
	opt_x = find_opt_val ("x", opts_found)
	opt_x = Trim(opt_x)
	if Left(opt_x,1) = "." Then opt_x = Mid (opt_x,2)
	fxtensions = Split (opt_x, ".")


	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	' Find filename
	' --------------------------------------------------------------------------------
	' - Start out with the filename variable
	' - If it does not have a filename extension add ".vbst"
	'		TODO: current error: a dot anywhere in an absolute fname would be seen as an extension
	' - If filename is absolute (ie. contains full path) use that and quit searching
	' - If file exists in current directory, use the corresponing absolute fname and quit searching
	' - Otherwise, check for filename in each directory in env variable PATH
	' - If found, use the corresponing absolute fname and quit searching
	' - If not found, search ends, and "absoluted" fname in current dir is used anyway,
	'	in which case a subsequent open will fail
	' - Open the absolute fname resulting from the above steps
	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	if InStr(filename, ".") = 0 then filename = filename & ".vbst"
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

	If Not opt_n And Not fso.FileExists(filename) Then filename = ""
		
	findFileName = filename

	Set fso = Nothing
	Set sh = Nothing
End Function '! Private Function findFileName(ByVal filename)

function psaylvarq (vname)
	psaylvarq = "say """ & vname &"='"" & " & vname & " & ""'"""
end function

function psaylvarcmq (vname)
	psaylvarcmq = ".say " & vname & "='$" & vname & "'"
end function

Private oRxSayLvar : Set oRxSayLvar = NewRegExp("(vname)", True, True)
function psaylvar2cmq (lvarname)
	psaylvar2cmq = oRxSayLvar.Replace(".say vname='$vname'", lvarname)
end function

function psaylvar3cmq (lvarname)
	psaylvar3cmq = Replace(".say vname='$vname'", "vname", lvarname)
end function

function psaywtype (lvarname)
	psaywtype = Replace("saywtype ""vname"", vname", "vname", lvarname)
end function

function b2s (b)
	b2s = "False" ' by default	
	If b Then b2s = "True"
end function

Sub saywtype (ByRef vname, ByRef value)
	Dim vtype, tname
	tname = typename (value)
	If 		tname = "String"  Then
		say vname & "(" & tname & ") = """ &     value  & """"
	ElseIf 	tname = "Boolean" Then
		say vname & "(" & tname & ") = " & b2s(value) & ""
	ElseIf  tname = "Integer" Then
		say vname & "(" & tname & ") = " & value
	ElseIf  vartype (value) = 9  Then
		say vname & " is an Object of type: " & tname
	Else
		say vname & "(" & t & ") = """ & value & """"
	End If
End Sub ' Sub saywtype (vname, ByRef value)


 ' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
 Private Sub RunTestFile (ByVal filename)
 ' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
	' RunTestFile - Similar to the "Import" routine, but some specific tricks apply:
	' ============================================================================
	' 
	' Preprocessing happens as follows for each line read from the file RunTestFile:
	'  0. "Normal lines": By default lines are executed as normal vbs code without change
	'		Lines are executed "immediately" ie. before the next line is read in from file
	'		However, different types of multi-line blocks are exceptions to this. See below.
	'
	'  1. "Dot substitution": Lines beginning with dot "." are converted as follows:
	'     Think of initial "." dot from a lazy-user perspective, such as: 
	'        Please run this command cmd ... for me, and btw I am too lazy to write punctuation, so
	'        could you please add quotes around whatever comes after the command, and while you're
	'        at it, I'd appreciate if you would substitute expressions within curly braces too.
	'     This translates into the following steps:  
	'     a. Double quotes existing in the line are doubled, from " to ""
	'     b. Expressions enclosed in curly braces {} are evaluated and substituted
	'        example: "linecount = {linecount}" the value of linecount is substituted
	'        or "Date: {date} Time: {time}" will make those substitutions
	'        NOTE: that ONLY GLOBAL variables are valid in such expressions
	'     c. Curly braces within quotes are not substituted
	'     d. Quotes in expressions within curly braces are not doubled, in fact they are executed,
	'        as part of evaluating the expression.
	'     e. Then, initial "." removed, and first token separated, the remaining line is enclosed
	'        in quotes and the resulting line executed:
	'        eg. The line: .say time={time} results in substitution of {time} with current time
	'     	 giving the line "say ""time=15:12:36""" which is then executed, giving
	'        the console output: time=15:12:36
	'		
	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	'  2. Slave marker: Lines beginning with the "slave marker" ">" or "_" are handled as follows:
	'     a. Quote-doubling and curly-brace substitution is done, same as for "dot-substitution" above
	'     b. Then, initial ">" or "_" removed, function ssend is called with the resulting line as argument
	'        eg. The line: _ -l setmem B7F8 "this is a string"
	'        becomes: ssend ("-l setmem B7F8 ""this is a string""")
	'		 Then, when executed, ssend will understand -l as a command line option, strip it off
	'     	 and send the line: 'setmem B7F8 "this is a string"' to stdin of the slave app
	'     	 (with the -l option, ssend "logs" ie. prints a message of what is sent just before sending)
	'     c. Note that the line sent on to the slave may well include an initial dot "."
	'     	 In case the slave-process is itself another instance of lookfor, then the line:
	'     	 "_.say hello" will send ".say hello" to the slave, which dot-converts it to 'say "hello"'
	'     	 and the text "hello" will be sent back in the slave's stdout
	'     
	'  3. # Hash-mark: Lines with an initial "#" are executed immediately, same as "normal" lines.
	'		Ie. these lines will work the same whether the hash-mark is there or not, including
	'		dot-substitution and slave-marker lines
	'		This is just to maintain consistency with "Hash-mark" lines within code blocks, see below.
	'		Think of these lines as similar to preprocessor directives and use them for that purpose.
	'		PLEASE DO NOT USE them for other "normal" code
	'		
	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	'  4. Code Blocks: Multi-line code blocks enclosed in "curly-brace" markup "{ <multi-line block> }"
	'	  a. All content enclosed within the curly-braces is read line-by line into memory
	'		 and not until reaching the end marker "}" is the whole block executed in one go.
	'		 this allows  multi-line constructs such as Sub(), Function(), If Then Else etc.
	'		 An exception to this are the "hash-mark" lines, see below.
	'	  b. "Dot-substitution" is done, as described above for lines with initial dot "."
	'		 But NOTE that the substitution happens at parse-time, which could cause unexpected
	'		 results inside Sub() or Function() definitions if these are called later because
	'		 the dot-substituted variables will still appear with the values they had at parse-time
	'		 eg. the line '.say Current time: {Time}' inside a Sub() will give the time at parse-time
	'	  c. NOTE: NOT done: Quote-doubling and curly-brace substitution is NOT done for normal lines
	'		 ONLY for lines with the initial dot
	'	  d. # Hash-mark: Lines with an initial "#" are similar to preprocessor directives, executed
	' 		 immediately at parse-time and NOT included as part of the block being read in.
	'		 However, a special syntax applies which is not fully documented yet.
	'	   d1: Note: Hash-mark lines are in global context and namespace, completely independent of
	'		 the surrounding lines of code in the block. Eg. local variables are not visible to them
	'	   d2: "dot-substitution" works in these lines too, as described above
	'	   d3: "slave-marker" lines also work "#_ <line>", but are unnecessary and NOT RECOMMENDED
	'	   d4: The #-substitution DOES make sense inside a block like: { Class xyz ... End Class  }
	'			 # If Global_xyzClass_defined Then GlobalDiscardThisBlock = True
	'		 Here the whole multi-line block will be discarded if the global var is set, eg.
	'		 to avoid an execution error by NOT repeating a class definition found in that block.
	'	  e. "Chatter Comments" Initial '!: Generates "Direct" .vbst parse-time comment output
	'		 equivalent to #.say <comment text> except that expression substitution is not done
	'		
	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	'  5. Here-Blocks: aka "here documents" Multi-line blocks of text enclosed in the markers "<+" and "+>"
	'	  An entire multi-line block of text is treated literally as one string, including linefeeds.
	'	  Useful for printing multi-line messages, or assigning multi-line text content to variables.
	'     a. Quote-doubling and curly-brace substitution is done, same as for "dot-substitution" above
	'	  b. The quote-doubling means that quotes appearing in the text will propagate correctly
	'		  eg. when being printed out, or read into a variable and printed out later.
	'	  c. Including literal curly braces is tricky at the moment:
	'		 The Chr function can be used: Left-curly as {chr(123)} and right-curly as {chr(125)}
	'			(actually right-curly "}" works as a normal character if not preceded by left-curly)
	'		 They can also be enclosed in "quotes {like this}", if you don't mind the quotes
	'		 TODO: Should probably implement an option for this, or handle backslash "\{" notation
	'		
	'  6. "Underscore-continuation": multi-line block of code,
	'	  signalled by space+underscore " _" at the end of each line, or just an underscore
	'	  in the case of empty lines.
	' 	  The input lines are concatenated into one long line, space-separated, without linefeeds, and
	'	  then executed, similar to multi-line blocks, but without dot- or other substitutions
	'	  PLEASE DO NOT USE this. It is cumbersome and doesn't seem to work right. Use multi-block instead
	'	  Left in for now, for possible compatibility issues.
	'		
	'		
	'!		@Todo Sub RunTestFile:Implement a state machine for handling line-by-line input 
	'!				making it possible to run also from the command line (good for testing)

	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	Dim exitOnError, fso, sh, file, code, dir, executeNow, nonExeCount
	Dim fLine, TestLine, qpos, ppos, semiLineNo, firstChar, withBlock, withLineCount

	filename = "  " & filename & "   "
	getopts ":E", filename, opts_found
	semiLineNo = 0
	withLevel		= 0			'!@withBlock 	 - > 0 while we are parsing inside (nested) "With ... End With" blocks
	withLineCount 	= 0			'!@withLineCount - Number of lines parsed in with-block

	exitOnError = False
	exitOnError = Not find_opt("E", opts_found)

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh  = CreateObject("WScript.Shell")
	executeNow	= True
	nonExeCount = 0
	fLine	 	= ""
	TestLine 	= ""
	code     	= ""


	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	' Find filename and open it
	' --------------------------------------------------------------------------------
	' - Start out with the filename variable
	' - If it does not have a filename extension add ".vbst"
	'		TODO: current error: a dot anywhere in an absolute fname would be seen as an extension
	' - If filename is absolute (ie. contains full path) use that and quit searching
	' - If file exists in current directory, use the corresponing absolute fname and quit searching
	' - If not, check for it in each directory in env variable PATH
	' - If found, use the corresponing absolute fname and quit searching
	' - If not found, search ends, and current dir "absoluted" fname is used anyway, and open will fail
	' - Open the absolute fname resulting from the above steps
	'saydbg "@all Initial filename="&filename
	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	if InStr(filename, ".") = 0 then filename = filename & ".vbst"
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
	saydbg "@runtestfile Opening file: " & filename
	
	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
	Set file = fso.OpenTextFile(filename, 1, False)
	flineno = 0

	' Main loop: Read from file, line by line
	' ----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+
	Do While Not file.AtEndOfStream 
		Do ' Dummy Loop to enable "Continue"
			code = ""
			semiLineNo = 0

			' >>>>>>>>>>>>>>>>>>>>> Read Input: Mode = Plain   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
			fLine 		= file.ReadLine		' An untrimmed fLine may be needed below in some cases
			flineno = flineno + 1
			blockLineCount = 1
			TestLine 	= rxLTrim(fLine)
			firstChar = Left(rxLTrim (Testline), 1)

			If rxRTrim (TestLine) =  "" Then Exit Do ' Exit Dummy Loop = "Continue"
			'If Left (TestLine, 1) = "'" Then Exit Do ' Exit Dummy Loop = "Continue"	' DISABLED, to enable specials in inital comments

			' Handle underscore-continued lines first. (So all the special cases below will work on the extended line)	
			' Note: Variable fline will hold the untrimmed, most recently read line from file.
			' 		TestLine will hold the combined line, with final underscores, newlines and intervening whitespace removed
			'!@Todo Todo: Handle underscore-continuations withing multi-line blocks
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			Do While Right(RxRTrim(TestLine), 2) = " _" Or Right(RxRTrim(TestLine), 3) = VBCrLf & "_"  Or RxTrim(TestLine) = "_"
				TestLine = RxRTrim (TestLine)
				TestLine = RxRTrim(Left(TestLine, Len(TestLine)-1))
				fLine = file.ReadLine
				TestLine = RxLTrim (TestLine & " " & RxTrim(fline))
				flineno = flineno + 1
				saydbg "@runtestfile-uscore Underscore-continue to line no: " & flineno & " TestLine:'" & TestLine &"'"
				If file.AtEndOfStream Then Exit Do
			Loop

			' "Percent-sign" lines: Function substitution at parse time
			'	eg. %myfunc arg - is converted to: myfunc ("arg") - which is then evaluated,
			'	and whatever was returned by myfunc is substituted on the line
			' 	Note: myfunc might well return a dot-line, to be further preprocessed in the next steps below
			'	Note: Not recommended, but it might even return a Hashmark-line or "Chatter comment"
			'	Note: Very much NOT recommended, WEIRD but interesting for experiments: it could return multi-line code
			'	Note: returning another percent-line FAILS (syntax error), as this code is not called recursively
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			If Left (TestLine, 1) = "%" Then
				saydbg "@runtestfile-percent Percent-sign preprocessing   initial: " & TestLine
				Testline = funcsubst (Testline)
				saydbg "@runtestfile-percent Percent-sign preproc after funcsubst: " & TestLine
				Testline = eval (Testline)
				saydbg "@runtestfile-percent Percent-sign preproc      after eval: " & TestLine
				TestLine = rxLTrim (Testline)		' Enable further dot- and comma-line processing below
			End If

			' Asterisk lines executed immediately, for debugging, WITHOUT dot or slave-marker -substitution 
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			If Left (TestLine, 1) = "*" Then						' **** DEBUG
				code = Mid(TestLine,2)
				'code = replace_args (code)
				'sayerr "**** Asterisk mode, executing: " & code
				executeglobal(code)
				Exit Do		' Continue
			End If

			' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
			' Hash-mark lines executed immediately, similar to preprocessor directives
			' TODO: Make hashmarking MORE like REAL preprocessor directives, especially allow include files.
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			If Left (TestLine, 1) = "#" Then			' #### DEBUG
				TestLine = Mid (rxLTrim(TestLine), 2)
				saydbg "@runtestfile Hashmark immediate Execute: " & TestLine
				RunVbshLine(Testline)
				Exit Do		' Continue
			End If

			' "Chatter Comments" Initial '!: Generates "Direct" .vbst parse-time comment output
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			If Instr(Testline, "'!:") = 1 Then			' #### DEBUG
				say Testline
				Exit Do		' Continue
			End If


			' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
			' Handle multi-line code block
			' ----+----+----+----+----+----+----+----+----+----L----+----+----+----+----+----+----+----+----+----C
			' ----.----o----.----o----.----o----.----o----.----L----+----o----+----o----+----o----+----o----+----|
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0

			' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
			' Handle Curly-brace-enclosed multi-line code block aka "curly-block"
			' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			' This is a multi-line block of code, enclosed by start and end tokens
			' that MUST be read in from file FIRST, and THEN executed as ONE CHUNK
			' The left and right curly-braces "{" and "}" are the start and end tokens
			'!@Todo: TestCase: Handle case where end token "}" is on same line
			'!@Todo: TestCase: what if fLine has end token followed by spaces(?)

			If Left (TestLine, 1) = "{" Then		' Left-curly block indicator 
				right_curly_found = False
				WithLevel = 0
			
				TestLine = Mid (TestLine, 2)
				Do	' Loop until end of current curly-brace block
					LTrimLine = RxLTrim (Testline)
					
					If oRxEscMark.Test (Testline) Then		' Handle initial "escape-mark lines"
															' ie. starting with special escape character: .,#€...

						'escMark = oRxEscMark.Replace (Testline, "$2")	' The escape-mark character that was found
						'escRest = oRxEscMark.Replace (Testline, "$3")	' Rest of line, after the escape-mark
						
						' "Percent-sign" lines: Function substitution at parse time
						'	eg. %myfunc arg - is converted to: myfunc ("arg") - which is then evaluated,
						'	and whatever was returned by myfunc is substituted on the line
						' 	Note: myfunc might well return a dot-line, to be further preprocessed in the next steps below
						'	Note: Not recommended, but it might even return a Hashmark-line or "Chatter comment"
						'	Note: Very much NOT recommended, WEIRD but interesting for experiments: it could return multi-line code
						'	Note: returning another percent-line would FAIL (syntax error), as this code is not called recursively
						' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
						If Left(LTrimLine,1) = "%" Then			' #### DEBUG
							TestLine = Mid (LTrimLine, 2)
							saydbg "@runtestfile-percent Percent-sign preprocessing   initial: " & TestLine
							Testline = funcsubst (Testline)
							saydbg "@runtestfile-percent Percent-sign preproc after funcsubst: " & TestLine
							Testline = eval (Testline)
							saydbg "@runtestfile-percent Percent-sign preproc      after eval: " & TestLine
							LTrimLine = rxLTrim (Testline)		' Enable further dot- and comma-line processing below
						End If

						' "Chatter Comments" Initial '!: Generates "Direct" .vbst parse-time comment output
						If Instr(LTrimLine, "'!:") = 1 Then			' "Chatter Comments" #### DEBUG
							say Testline
							Testline = ""
						End If

						' "Hashmark" lines: Initial "#" is executed immediately at parse-time
						If Left(LTrimLine,1) = "#" Then			' #### DEBUG
							TestLine = Mid (LTrimLine, 2)
							saydbg "@runtestfile hashmark immediate Execute: " & TestLine
							RunVbshLine(Testline)
							Testline = ""
						End If

						If Left(LTrimLine,1) = "." Then		' Preprocess dot-lines
							If Left(LTrimLine,2) = ". " Or withLevel <= 0 Then		' Inside a With-block the dot must be followed by a space
								Testline = cmdsubst (TestLine)
								saydbg "@runtestfile-dot dot-preprocessed result:"&TestLine  ' **** DEBUG
							End If
						End If

						If Left(LTrimLine,1) = "," Then		' Preprocess comma-lines
							saydbg "@runtestfile-dot dot-preprocessing:"&TestLine  ' **** DEBUG
							Testline = preprocess_cmdline (TestLine)
							saydbg "@runtestfile-dot dot-preprocessed result:"&TestLine  ' **** DEBUG
						End If

					End If	' If oRxEscMark.Test (Testline) Then		' Handle initial "escape-mark lines"
					
					saydbg "@runtestfile-curly curly block line: " & TestLine
					
					'!@ TODO: Consider RECOGNIZING closing-curly at the END of a trailing comment

					' Detect "With ... End With" blocks
					' NOTE: "With " and "End With" must appear FIRST on the line, one space after "With" is required
					'!@Todo	Todo: Error if curly-block is closed without ending a with-block
					If InStr(1, LTrimLine,"With ", 1) = 1 Then
						saydbg "@runtestfile-with With detected:"&TestLine
						If withLevel = 0 Then withLineCount = 0
						withLevel = withLevel + 1
					End If

					' Count number of lines in with-block (including nested blocks if any)
					' Note: the count is not reset at end of with-block, only on start of the next one.
					'		ie. understand withLineCount as "number of lines found in last with-block"
					'		unchanged until the next one is found
					If withLevel > 0 Then withLineCount = withLineCount + 1	

					' Detect End of a with-block
					If withLevel > 0 And InStr(1, LTrimLine,"End With", 1) = 1 Then
						saydbg "@runtestfile-with With ended:"
						withLevel = withLevel - 1
					End If

					If InStr(Testline,";") > 0 Then			' Generate helpful error info (line number) if semicolon used
						If Right(RxRTrim(unComment(Testline)), 1) = ";" Then
							sayerr "(RunTestFile): Warning: Eliminating final semicolon in file: " & shortfilename & ", Line no: " & flineNo
							Testline = RxRTrim(unComment(Testline))
							Testline = Left(TestLine, len(TestLine) - 1)
						ElseIf InStr(checkQ(Testline),";") > 0 Then
							semiLineNo = flineNo
							sayerr "(RunTestFile): Warning: Unable to eliminate semicolon in file: " & shortfilename & ", Line no: " & flineNo
							sayerr "Line: "&TestLine
						End If
					End If


					' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
					If InStr(Testline,"}") > 0 Then			' Check for end of block
						' Check if there is an end token "}" anywhere in the line: if so, it is a possible candidate for end of block
						' (or the token could be inside a quote or part of a curly-brace substitution expression)
						' Hopefully it is more efficient this way as we avoid calling checkQ(Testline) for every line
						' Anyway: Now we DID find end token, so we call checkQ to see it is a real or a false alarm
						If Right(rxRTrim(checkQ(Testline)), 1) = "}" Then		' Note: CheckQ removes "regular" curly-brace pairs first
							' End token "}" found as last char, after comments and trailing spaces removed
							' ALSO: ensuring that previous {curly-brace enclosed} elements on the same line are disregarded
							' NOTE: It is important to call checkQ above, and then unComment below
							' Note for the future: End token "}", trailing spaces and comment are removed from Testline
							' ie. any preprocessing tags in trailing comments are lost beyond this point (if implemented in the future)
							' Should not matter: as it is strongly recommended to have the multi-line end token "}" on its own anyway
							' (on a separate line). Such a line _can_ have a trailing comment but any preprocessing tags included
							' in that comment will be lost, unless handled by logic above this point in the code.
							Testline = RxRTrim(unComment(Testline))
							Testline = Left(TestLine, len (TestLine) - 1 )
							right_curly_found = True
						ElseIf InStr(checkQ(Testline),"}") > 0 Then
							sayerr "(RunTestFile): Warning, in file: " & shortfilename & ", Line no: " & flineNo
							sayerr "Code after close token } not allowed. Code on following line discarded: "&TestLine
							TestLine = ""
							right_curly_found = True
						End If
					End If

					' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
					If len (code) > 0 Then code = code & VBCrLf 
					code = code & TestLine

					' Exit the loop if finished
					If right_curly_found Or file.AtEndOfStream Then Exit Do ' Finished reading current curly-brace block

					' Read next line within the block
					' >>>>>>>>>>>>>>>>>>>>> Read Input: Mode = Curly-Multiline   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
					fline = file.ReadLine
					flineno = flineno + 1
					blockLineCount = blockLineCount + 1
					TestLine = fline
				Loop ' ' Loop until end of current curly-brace block
				' saydbg "@runtestfile curly block found:" & vbcrlf & "----------" & vbcrlf & code & vbcrlf & "----------"

				'!@TODO: The discarding of blocks should be implemented as preprocessor directives instead
				'! (ie. #define blockxyz ... #ifdef / #ifndef ...)
				'! Drawback: Defined symbols must be kept as global vars, so it really boils down to the same thing
				'! Current solution is just a durty hack: escape-mark lines are unaffected, ie. currently not discarded
				If GlobalDiscardThisBlock Then
					saydbg "@runtestfile Discarding THIS block {" & VBCrLf & code & VBCrLf & "}"
					code = ""
					GlobalDiscardThisBlock = False
				End If
				If GlobalDiscardNextBlock Then
					saydbg "@runtestfile Discarding this NEXT block {" & VBCrLf & code & VBCrLf & "}"
					code = ""
					GlobalDiscardNextBlock = False
				End If
				TestLine = ""
			End If	' If Left (TestLine, 1) = "{" Then		' Left-curly block indicator 


			' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
			' Handle "here-block" (aka "here document")
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			' The "here-block" is a multi-line block of text, enclosed in "<+" and "+>" respectively
			ppos = InStr (TestLine, "<+")
			If ppos > 0 Then		' Left <+ here-block indicator 
				'saydbg "@runtestfile_here TestLine="&TestLine
				'saydbg "@runtestfile_here ppos="&ppos
				' Check for string literal preceding the start token:
				' Such a str will be concatted to, but substitution does not happen
				
				' Check number of " double-quote chars appearing before the start token
				qcount = len(Left(TestLine, ppos)) - len(replace(Left(TestLine, ppos), """", ""))
				' Odd number of quotes:the start token <+ is INVALID and DISREGARDED, as it is part of a string literal
				' Even number or zero quotes: start token <+ is valid
				
				' If comment marker found earlier on line, DISREGARD start token as it is part of a comment
				' This is accomplished by "pretending" that there was exactly ONE preceding quote char
				'TODO: This check is NOT foolproof: Single quote inside a double-quoted sequence would count as comment
				If InStrRev(Left(TestLine, ppos), "'") > 0 Then qcount = 1 ' (here we pretend the count was 1)

				'saydbg "@runtestfile_here qcount="&qcount
			
				' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
				If qcount mod 2 = 0 Then
					'saydbg "@runtestfile_here Even number of quotes detected, or none: Thus, a VALID here-block start marker was found"

					code    = Left(TestLine, ppos - 1)
					remline = Mid (TestLine, ppos + 2)

					if len(remline) > 0 then
						code = code & """" & replace_args (remline) & """ & VBCrLf"
					Else
						code = code & """"""
					End If 
					'saydbg "@runtestfile_here code="&code


					Do ' Loop until end of current here-document
						' >>>>>>>>>>>>>>>>>>>>> Read Input: Mode = Here-block   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
						fLine = file.ReadLine
						flineno = flineno + 1
						blockLineCount = blockLineCount + 1
						TestLine = fLine

						If Right(RTrim(TestLine), 2) = "+>" Then
							TestLine = RTrim(TestLine)
							TestLine = Left(TestLine, len (TestLine) - 2 )
							code = code & " & """ & replace_args (TestLine) & """"
							Exit Do
						Else 
							code = code & " & """ & replace_args (TestLine) & """ & VBCrLf"
						End If
						If file.AtEndOfStream Then Exit Do ' Loop until end of current here-document
					Loop ' Loop until end of current here-document
					'saydbg "@runtestfile_here here block found:" & vbcrlf & "----------" & vbcrlf & code & vbcrlf & "----------"
					TestLine = ""
				End If
			End If	' If ppos > 0 Then		' Left <+ here block indicator 

			' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
			' Handle "Underscore-continued" blocks
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			' multi-line block of code, signalled by space+underscore " _" at the end of each line,
			' or just an underscore in the case of empty lines
			' Note: The input lines are concatenated into one long line, space-separated, without any linefeeds
			' TODO: Analyse possible conflict between Underscore-continued lines and multi-line blocks
			if len (code) = 0 Then code = TestLine
			' Do While Right(code, 2) = " _" Or Right(code, 3) = VBCrLf & "_"  Or code = "_"
			Do While Right(RxRTrim(code), 2) = " _" Or Right(code, 3) = VBCrLf & "_"  Or RxTrim(code) = "_"
				' >>>>>>>>>>>>>>>>>>>>> Read Input: Mode = Line-continuation   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
				code = RxRTrim (code)
				code = RTrim(Left(code, Len(code)-1)) & " " & RxTrim(file.ReadLine)
				flineno = flineno + 1
				blockLineCount = blockLineCount + 1
			Loop

			' Handle lines to be piped to the stdin of a slave process
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			' Lines beginning with a slave marker "_" or ">" are sent to stdin of slave app
			' In order to be sent off correctly by the ssend function, the lines are preprocessed
			' with quote-doubling, {} curly-brace substitution and quote-enclosed
			' NOTE: Syntax for command line options eg. "_ -lE <rest of line>"
			' 		where ssend takes -lE as options and <rest of line> is sent to the slave's stdin
			' TODO: if the slave itself is another lookfor instance, it should refuse these lines here.
			'		Slave-of-slave or multi-slave scenarios don't seem worth supporting for now.
			' TODO: TestCase: slave-sending of multiline code, underscore-continued or { multi-line } code blocks
			If Left (code, 1) = "_" Or Left (code, 1) = ">" Then
				code = Mid(code,2)
				code = replace_args (code)
				'code = Replace(code,"""","""""")
				
				'saydbg "@runtestfile_ssend Test output: ssend(" & code & ")"		' *** DEBUG
				'code = "ssend(" & code & ")"
				
				code = "ssend(""" & code & """)"
			End If
		

			' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
			' SPOX (Single Point of Execution): Execute the line, or block of lines that were read in
			' Note: except for "asterisk lines" for debugging that are executed above
			' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
			If executeNow Then
				' Boolean executenow is for handling  multiline input, currently always True ie. NOT USED
				saydbg "@runtestfile Running code: " & code
				nonExeCount = 0
				MyErr.Reset
				RunVbshLine(code)

				If MyErr.Number <> 0 Then
					If semiLineNo > 0 Then
						sayerr "(RunTestFile): Execution Error: Check invalid semicolon in file: " & shortfilename & ", Line no: " & semiLineNo
					ElseIf withLineCount > 0 Then
						sayerr "(RunTestFile): Execution Error: Expected ""End With"". Check ""With""-statement in file: " _
						& shortfilename & ", Line no: " & withLineNo
					Else 
						sayerr "(RunTestFile): Execution Error in file: " & shortfilename & ", Line no: " & flineNo
						'!@TODO: Limit the output of code here to the first n lines
						sayerr "Code that failed: "&code
					End If
					If exitOnError Then Exit Sub
				End If

			ElseIf Not executeNow Then		' NOTE: Currently NOT used 
				nonExeCount = nonExeCount+1
			End If
		Loop While False ' Dummy Loop to enable "Continue"
	' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
	Loop ' Do While Not file.AtEndOfStream 

	file.Close()
	Set fso = Nothing
	Set sh  = Nothing

  	' ----+----1--  RunTestFile  --3----+----4----+----5--  RunTestFile  --7----+----8--  RunTestFile  --O
 End Sub '! Private Sub RunTestFile(ByVal filename)

 sub rtf(s): RunTestFile s: end sub		' *** DEBUG

_ 	
 Function sans_prompt (s, prompt)
	sans_prompt = Replace (s, vbCrLf & prompt, "") '! Not sure if we should remove only one here (?)
	'!sans_prompt = Replace (s, prompt) '! Not sure if we should remove these (?)
 End Function '! Function sans_prompt (s, prompt)

 Function checkQ (ByRef s)
 ' Return a string with quotes, curlies and comments removed.
 ' eg. the line: dosomething with "{this} 'string'" and {a "variable"} hello {another var}' a comment
 ' becomes: dosomething with  and  hello 
 	Dim pos1, pos2, pos3, remain 
	 checkQ = ""
	 remain = s
	 Do While len (remain) > 0
		 saydbg "@checkq remain="&remain
		 pos1 = InStr (remain, """")
		 pos2 = InStr (remain, "{")
		 pos3 = InStr (remain, "'")

		If pos3 > 0 and (pos3 < pos1 or pos1 = 0) and (pos3 < pos1 or pos1 = 0) Then		' A comment stops the game right here
			checkQ = checkQ & Left(remain, pos3 - 1)
			exit function

		ElseIf pos1 > 0 and ( pos1 < pos2 or pos2 = 0 ) then		' Double-quote char found
			checkQ = checkQ & Left(remain, pos1 - 1)
			remain = Mid (remain, pos1 + 1)
			endpos = InStr(remain, """")

		ElseIf pos2 > 0 and ( pos2 < pos1 or pos1 = 0 ) then		' left curly-brace found
			checkQ = checkQ & Left(remain, pos2 - 1)
			remain = Mid (remain, pos2 + 1)
			' Some added sophistication for detecting where a curly-brace-enclosed element "really" ends
			'endpos = InStr(remain, "}")				' This usually works, but fails for curly-within-quote-within-curly
			endpos = len (getUntilEnd (remain, "}"))	' This one works, even for ... {"str{curly}"} ..., but NO deeper than that
														' This also works: ...{say "end token is ""}"""}
		
		Else
			checkQ = checkQ & remain				' nothing found: no comments, curly-braces or quotes
			endpos = 0
		End If

		 if endpos = 0 Then Exit Function			' nothing found, or closing quote or curly-brace missing

		 remain = Mid (remain, endpos + 1)
	 Loop
 End Function  'Function checkQ (ByRef s)

 Function unComment (ByRef s)
 ' Return a string with final comment removed.
 ' eg. the line: whatever "this is 'a string'" and { another 'one} again ' comment text
 ' becomes: whatever "this is 'a string'" and { another 'one} again 
 ' but single-quotes within quotes or within curly-braces are unaffected
 	Dim pos1, pos2, pos3, endpos, remain 
	If Not InStr (s, "'") > 0 Then
		unComment = s
		Exit Function
	End If
	
	unComment = ""
	remain = s
	Do While len (remain) > 0
		saydbg "@uncomment remain="&remain
		pos1 = InStr (remain, """")
		pos2 = InStr (remain, "{")
		pos3 = InStr (remain, "'")

		If pos3 > 0 and (pos3 < pos1 or pos1 = 0) and (pos3 < pos1 or pos1 = 0) Then		' A comment stops the game right here
			unComment = unComment & Left(remain, pos3 - 1)
			exit function

		ElseIf pos1 > 0 and ( pos1 < pos2 or pos2 = 0 ) Then		' Double-quote char found
			unComment = unComment & Left(remain, pos1)
			remain = Mid (remain, pos1 + 1)
			endpos = InStr(remain, """")
			unComment = unComment & Left(remain, endpos)

		ElseIf pos2 > 0 and ( pos2 < pos1 or pos1 = 0 ) Then		' left curly-brace found
			unComment = unComment & Left(remain, pos2)
			remain = Mid (remain, pos2 + 1)
			endpos = InStr(remain, "}")
			unComment = unComment & Left(remain, endpos)

		Else
			endpos = 0
		End If

		if endpos = 0 Then
			unComment = unComment & remain
			Exit Function			' closing quote or curly-brace missing
		End If

		 remain = Mid (remain, endpos + 1)
	 Loop
 End Function  ' Function unComment (ByRef s)

 Set gend = getref("getUntilEnd")

 Function getUntilEnd (ByRef s, endchar)
 	' Extract from beginning of string until end character is found, skipping characters within quotes
	 Dim c, pos1, pos2, endposq, remain, result
	 getUntilEnd = ""
	 result = ""
	 c = Left(endchar,1)
	 If c = "" Then Exit Function
	 saydbg "@getuntilend c="&c

	 If InStr (s, c) = 0 Then		' endchar not found: Nothing to do
		saydbg "@getuntilend c NOT FOUND"
		getUntilEnd = s
		Exit Function
	 End If

	remain = s
	Do While len (remain) > 0
		saydbg "@getuntilend remain="&remain
		pos1 = InStr (remain, """")
		pos2 = InStr (remain, c)

		If pos1 > 0 and ( pos1 < pos2 or pos2 = 0 ) Then		' Double-quote char found
			result = result & Left(remain, pos1)
			remain = Mid (remain, pos1 + 1)
			endposq = InStr(remain, """")
			result = result & Left(remain, endposq)
			remain = Mid (remain, endposq + 1)

		ElseIf pos2 > 0 and ( pos2 < pos1 or pos1 = 0 ) Then		' endchar found
			result = result & Left(remain, pos2)
			Exit Do

		Else														' endchar not found in remain
			result = result & remain
			Exit Do
		End If

	Loop
	getUntilEnd = result
 End Function ' Function getUntilEnd (ByRef s, endchar)




_
 '  SlveReadUpto (pattern)
 ' Name alternatives:
 ' ReadChild (pattern)
 ' GetUpto (pattern)
 ' GetSlaveUpto (pattern)
 ' GetSlvUpto (pattern)
 ' SlvGetUpto (pattern)
 ' SlaveGetUpto (pattern)
 ' SlveReadUpto (pattern)	- Chosen for now, also: SlveWrite[ln], SlveRead[ln], NewSlve etc.
 ' SlaveReadUpto (pattern)
 ' ChldReadUpto (pattern)
 ' SlaveReadUpto (pattern)
 ' SlveReadUpto (pattern)
 ' ====================================================================================================
 ' Read output from Slave app's stdout, until a stop pattern is found.
 ' Return value: All text found up to and including the stop pattern
 ' NOTE: The pattern MUST be found in the output buffer, if not it WILL hang forever
 '		Likewise, we MUST NOT read even one character past the end of the pattern.
 ' Normally you'll want to use the slave app's command prompt as pattern, in which
 ' case this routine reads all output up to, and including, the next prompt.
 ' If you are lucky enough that the prompt is configurable, you can set it to something
 ' distinctive like "==:)>> " which is unlikely to be found anywhere else in the output.
 ' NOTE: If a stop pattern occurs in output before the next prompt, this leads to slave
 '		output being out of sync with input, ie. a fault situation that will mix things
 '		up, but goes completely undetected here. It is squarely up to the calling
 '		logic to ensure this does not happen or to handle it if it does.
 ' TODO: Create a parallell pattern matching routine with regexp's and benchmark
function SlveReadUpto (pattern)
	s = ""
	p = 0   				'! p > 0 indicates pattern found
	a = left(pattern,1)   	'! a = first char in pattern
	z = right(pattern,1)  	'! z = last char in pattern
	'y = right(pattern,2)
	'y = left (y,1)  		'! y = last char but one in pattern (not used)
	a0 = 0 					'! a0 = candidate position of beginning of pattern when a found
	ap = 0 					'! ap = subsequent candidate position of beginning of pattern
	afound = false
	LastVbCr = 0
	
	'say "a=/"&a&"/"
	'say "z=/"&z&"/"
	'say "afound="&afound
	'say "pattern=/"&pattern&"/"
	do while (p <= 0)
		c = SlaveExec.stdout.read(1)

		s = s & c
		'!say "**** len(s)="&len(s)&" s=/"&s&"/"
		If  c = VbCr then LastVbCr = len(s) - 1



		'! Quickly Skip this part until a found
		if afound then 
			if c = a then ap = len (s)
			'! if c = z then p = instr(1,mid(s, a0),pattern)
			if c = z then
				p = InStr(a0,s,pattern) 

			'! Decide if a should start over as "unfound" again to speed things up.
			'! This only happens on finding a sequence, longer than pattern, where
			'! neither a nor z is included ie. finding pattern there is impossible.
			elseif len (s) - ap > len (pattern) then
				'! say "**** resetting afound at pos="&len(s)&" str=/"& mid(s,ap)&"/"
				afound = false
				a0 = 0
				ap = 0
			end if

		'! Quickly search for the first a
		elseif c = a then
			afound = true
			a0 = len (s)
			ap = len (s)
			'! say "**** found a0 at pos="&len(s)&" s=/"& s &"/"
			if len(pattern) = 1 then p = len(s) '! handles pattern of length 1
		end if
	loop

	if HideSlavePrompt then
		s=left(s,len(s)-len(pattern))
		if right(s,1) = vbLf then s = left(s,len(s)-1)
		if right(s,1) = vbCr then s = left(s,len(s)-1)
		'if (a0 - LastVbCr = 2) and (LastVbCr > 0) then
		'	s = left(s,LastVbCr - 1)
		'end if
	End If

	's=trim(s)
	'if len(s) = 0 then sayerr "**** (SlveReadUpto): found empty string"			' **** DEBUG
	'if s = vbCrLf then sayerr "**** (SlveReadUpto): found single vbCrLf"			' **** DEBUG

	SlveReadUpto = s
	
end function '! function SlveReadUpto (pattern)

 ' ====================================================================================================
 '_h1 The 100 Hundred Doors Problem
 '_<h1> The 100 Hundred Doors Problem // "formal" notation for h1 (simplified above works too)
 ' ====================================================================================================
_
 '_TODO: Transfer this to some other file
 ' Explanation (from somewhere on the net):
 ' A 100 hundred doors are initially closed
 ' Then you get busy going from door 1 up to 100 and toggle each of them closed
 ' Then start at 2 and toggle every second door
 ' Then from 3 toggle every third
 ' etc. up to 100. - what doors are open at the end?
 ' whichopen_doors(100) below will tell you, for 100 doors, or n doors
 comment_="" '_h1 The 100 Hundred Doors Problem
 redim doors(100)
 sub ini_doors(n)
 	redim doors(n)
 	for i = 1 to n
 		doors(i) = 1 ' 1 means closed
 	next
 end sub
 
 sub toggle_doors(n,skip)
 	i = skip
 	do while i <= n
 		doors(i)= 1 - doors(i)
 		i = i + skip
 	loop
 end sub
 
 sub say_doors(n)
 	s_doors = " "
 	j=1
 	k=1
 	for i = 1 to n
 		if k = 1 then s_doors = s_doors & right("      "&i,6) & ": "
 		s_doors = s_doors & doors(i)
 		if j >= 10 then
 			if k>=50 then
 				s_doors = s_doors & vbCrLf
 				k=0
 			end if
 			s_doors = s_doors & " "
 			j = 0
 		end if
 		j = j+1
 		k = k+1
 	next
 	say s_doors
 end sub
 
 sub toggle_alldoors(n)
	 ini_doors n
	 stepsize = 10 * int (n/100)
	 i_shown = 0
 	say "Start toggling with doors("&ubound(doors)&")"
 	Dim i
 
 	j=1
 	for i = 1 to n
 		toggle_doors n, i
	'	if j >= 10 then
 		if j >= stepsize then
 			say "Toggled up to " & i
 			say_doors n
			i_shown = i
 			j = 0
 		end if
 		j = j+1
 	next
	If i > i_shown + 1 Then
		say "Toggled up to " & i - 1
		say_doors i-1
 End If
 
 	say "Finished toggling at " & i
 
 end sub

 sub whichopen_doors()
 	s_doors = ""
 	for i = 1 to ubound(doors)
 		if doors(i) = 0 then
 			s_doors = s_doors & " " & i
 		end if 
 	next
 	say s_doors
 end sub
 
 ' =============================================================================
 '_/h1 The 100 Hundred Doors Problem
 ' =============================================================================
 
_
 ' ====================================================================================================
 '_h1 Obsolete, Old, Deprecated and Forgotten
 ' Old stuff kept around in here just in case, to be thrown out soon
 ' ====================================================================================================
 
 ' DELETED 2018-06-30: Private Sub old_RunTestFile(ByVal filename)

' 2018-08-18: OLD - readToken0 (ByRef s) - Get first token from a command line
' ====================================================================================================
Function readToken0 (ByRef s)
	readToken0 = ""
	Dim t
	t = ""
	Do While t = ""
		arr = Split (s, " ", 2)
		If ubound(arr) > 0 Then
			t = arr(0)
			s = arr(1)
		ElseIf ubound(arr) = 0 Then
			t = arr(0)
			s = ""
		Else
			s = ""
			Exit Do
		End If
	Loop
	readToken0 = t
End Function ' Function readToken0 (s)



 ' ====================================================================================================
 '_h1 Includefiles that must be run at the end
 ' ====================================================================================================
 ' ====================================================================================================

'RunTestFile ""

RunTestFile "TestCase.vbst"
RunTestFile "misc.vbst"

