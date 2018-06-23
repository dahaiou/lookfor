
' Global Declarations
' =================================================================================================
IncludeName = "inc1.vbs"
IncludeVersion = "0.06-01"
inc1NameString = IncludeName & " V" & IncludeVersion
say "Including: " & inc1NameString

ProgPackages = ProgPackages & " " & Include_blurb

Dim g_opts_found
g_opts_found=""

Dim oRx, oMatch

 '! Global Regex to use anywhere
	Set oRx = New RegExp
	oRx.global = True
	oRx.ignorecase = False
	oRx.pattern = "^(\s*)"
	Set oMatch = oRx.Execute("Dummy text")



' Routines
' =================================================================================================

'! Prepare a new regular expression.
'! By Ansgar Wiechers, See Copyrights Ref.1
'! Name changed from CompileRegExp to NewRegExp
'!
'! @param  pattern      The pattern for the regular expression.
'! @param  ignoreCase   Boolean value indicating whether the regular expression
'!                      should be treated case-insensitive or not.
'! @param  searchGlobal Boolean value indicating whether all matches or just
'!                      the first one should be returned.
'! @return A new regular expression object.
Private Function NewRegExp(pattern, ignoreCase, searchGlobal)
	Set NewRegExp = New RegExp
	NewRegExp.Pattern    = pattern
	NewRegExp.IgnoreCase = Not Not ignoreCase
	NewRegExp.Global     = Not Not searchGlobal
End Function



Sub ssend (ByRef cmdline)
	'! Usage: ssend -E -l <line>
	'!@param -E - Echo off (on is default) Do not echo return text from slave app (but read into FoundLine)
	'!@param -l - Logging on (off is default) Echo <line> just before sending it
	Dim opts_found		' Make sure opts_found is local
	Dim opt_echo
	opts_found = ""
	saydbg "@ssend calling getopts with :El " & cmdline
	getopts ":El", cmdline, opts_found

	opt_echo = true
	if find_opt("E", opts_found) then opt_echo = false
	
	if SlaveExec Is Nothing then
		sayerr "Error: No slave app to send to."
		Exit Sub
	End If
	saydbg "@ssend calling find_opt ( l, " & opts_found
	if find_opt("l", opts_found) Then TClog "(ssend) Sending: " & cmdline
	
	saydbg "@ssend Sending: " & cmdline
	SlaveExec.StdIn.WriteLine(cmdline)
	FoundLine = SlveReadUpto (SlavePrompt)
	if opt_echo And len(FoundLine) > 0 Then
		'say SlaveOutFlag & FoundLine
		'say "with regex:"
		oRx.global = true
		' oRx.ignorecase = false		' not needed
		oRx.pattern = "\n"
		TClog SlaveOutFlag & oRx.Replace (FoundLine, VBCrLf & SlaveOutFlag)
	End If

End Sub ' ssend(cmdline)


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
 ' ====================================================================================================
 ' Globals related to saydbg(s)
 Dim DBG_enabled, DBG_current, DBG_banner
 DBG_enabled = "|"		' holds enabled debug topics separated by vertical bar eg. "|topic1|topic2|...""
 DBG_topic   = "|"
 DBG_banner  = "#### DEBUG"


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
		sayvarq "DBG_enabled"
		sayvarq "DBG_topic"
		sayvarq "DBG_banner"
		

	Elseif sDBcheck ("|reset|", cmd) Then
		' reset
		say "dbg reset:"
		say "dbg reset: "
		sDBreset DBG_enabled


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

 Sub saydbg (s)
	' TODO: Selective debug messages: enabled/disabled by topic or function
	' Global vars:  DBG_enabled = /abc/def/regex/...  these ones are enabled
	'				DBG_current = /regex/ - current topic, ie, must be found
	'				alternatively, the topic can be passed in as parameter
	'				signalled by initial at-sign dash eg. saydbg "@regex <message> "
	Dim topic, topics, t, on_topic
	topic = ""
	' Get topic from s, if present, in the form "@topic rest of string"
	If InStr (LTrim(s),"@") = 1 Then
		sv    = split(LTrim(s), " ", 2)
		topic = Trim(Replace(sv(0),"@", "", 1, 1))
		if ubound(sv) >= 1 Then
			s = sv(1)
		Else
			s = ""
		End If
	End If
	topics = "|" & topic & "|" & DBG_topic 
	' TODO: Decide: Should empty s be rejected here ? Currently it's not
	
	' Now check if we are on-topic
	If sDBcheck (DBG_enabled, "all") Then		' "all" globally enabled means we are always on-topic
		'WScript.Stderr.WriteLine DBG_banner & "(saydbg): ALL enabled: Forced on_topic"
		on_topic = True
	Else
		on_topic = False ' By default
		For Each t in Split(topics, "|")
			t = Trim (t) ' Probably not necessary but seems right
			If len(t) > 0 And sDBcheck (DBG_enabled, t) Then
				'WScript.Stderr.WriteLine DBG_banner & "(saydbg): on_topic enabled, key: " & t
				on_topic = True
				Exit For
			End If
		Next
	End If

	' If Not on_topic Then WScript.StdErr.WriteLine DBG_banner & "(saydbg)  NOT on topic: " & s
	
	If Not on_topic Then Exit Sub
	
	bann = DBG_banner & ": "
	If len (topic) > 0 Then bann = DBG_banner & "(" & topic & "): "
	
	
	WScript.StdErr.WriteLine bann & s
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


' readToken (ByRef s) - Read and remove first token from a command line
' ====================================================================================================
Function readToken (ByRef s)
	readToken = ""
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
	readToken = t
 End Function ' Function readToken (s)

'_h2 Sub ListProcessRunning()
' ====================================================================================================
_ 

'_h2 Test RegEx
' ====================================================================================================
_ 
 
 sub test_rx(s,sPat)
 	set oRx=Nothing
 	set oMatch=Nothing
 
 	Set oRx = New RegExp
 	oRx.global = true
 	oRx.ignorecase = false
 
 	oRx.pattern = sPat
 	Set oMatch = oRx.Execute(s)
 
 	sayomatch(oMatch)
 		
 end sub
 
 Sub sayomatch (Match)
 	say "  Match.Count=" & Match.Count
 	If Match.Count <= 0 Then Exit Sub
 
 	For i = 0 to oMatch.Count -1
 		say "    Match("&i&").submatches.Count=" & Match(i).submatches.Count
 		For j = 0 to Match(i).submatches.Count - 1
 			say "      oMatch("&i&").submatches("&j&")=" & oMatch(i).submatches(j)
 		Next
 	Next
 End Sub
  


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
	optstr = "abcde:f:x:z|good-times|parse-slowly:|a-good-opt:|"
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

'_h2 function split_optarr (s) - After getopt, split options into an array
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
	oRegOpts.pattern = "^(\s*[^-\s][\S]*[\s]*)?"				' Regex to remove initial command token, if any
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
		'say "**** DEBUG: opts_found:" & opts_found			' **** DEBUG
		'say "**** DEBUG: Checking:" & oRegOpts.pattern			' **** DEBUG
		'say "**** DEBUG: Against :" & restline			' **** DEBUG
 
		Set oMatch = oRegOpts.Execute(restline)
		
		'say "**** DEBUG: oMatch.Count=" & oMatch.Count & " oMatch.submatches.Count=" & oMatch.submatches.Count
		'say "**** DEBUG: oMatch.Count=" & oMatch.Count
		If oMatch.Count <= 0 Then Exit Do

		'say "**** DEBUG: oMatch(0).submatches.Count=" & oMatch(0).submatches.Count

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
						restline = this_opt & LTrim(restline)		' stick it in before rest of line
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
	
	'saydbg "@getopts Result: ========================================="
	'saydbg "@getopts opts_found  :" & opts_found
	'saydbg "@getopts Rest of line:" & restline
	
	If len(opts_found) < 2 Then 
		' Do nothing: No opts found, so DON'T CHANGE cmdline AT ALL
	ElseIf len(cmd_token) > 0 Then 
		cmdline = cmd_token & " " & restline
	Else
		cmdline = restline
	End If
	saydbg "@getopts resulting cmdline:'" & cmdline & "'"

End Sub ' Sub getopts (optstr, ByRef cmdline, ByRef opts_found)


' =================================================================================================
Private Sub RunTest(ByVal filename)
	' take filename as input
	'		extension .vbst is added unless given
	'		other extensions (dot anywhere in filename) are preserved
	'		filename. overrides: dot is removed and no extension used
	' log file is filename.vbstlog
	'		TODO: nifty way to allow other extensions
	'		Initial test stamp logged: Date filename etc. 
	'		previous logfile is overwritten
	'		TODO: option to rename previous logfile
	'		TODO: option to append to logfile instead
	'		logfile is kept open through to the end of execution
	'		log output is flushed continuously on every write
	'		TODO: option to append to logfile and close on every write
	'				(avoids losing log output on deadlock)
	'		TODO: options for summary/normal/detailed log output
	'				detailed: 	- Initial presentation blurb for each test case
	'							- "comment" log commands eg. "setting up for test xyz"
	'							- all sent setup commands are logged (eg. setting memory)
	'							- all setup output received back (confirmations activated if present)
	'							- sent test commands are logged
	'							- stepwise execution with logging defined at each step
	'							- "misc" detailed schemes yet to be thought of
	'							- all test output received back
	'							- all test results 
	'							- all output/lookfor comparisons on failure
	'				normal: 	- Initial presentation blurb for each test case
	'							- "comment" log commands eg. "setting up for test xyz"
	'							- NOT all sent setup commands are logged (eg. setting memory)
	'							- NOT all setup output received back (confirmations activated if present)
	'							- all sent test commands are logged
	'							- NOT (probably) stepwise execution with logging defined at each step
	'							- NOT (probably) "misc" detailed schemes yet to be thought of
	'							- all test output received back
	'							- all test results 
	'							- all output/lookfor comparisons on failure
	'				summary: 	- NO Initial presentation blurb for each test case
	'							- Possibly test section logs ie. log starting in on new chapter 
	'							- NOT all "comment" log commands 
	'							- NOT all sent setup commands are logged (eg. setting memory)
	'							- NOT all setup output received back (confirmations activated if present)
	'							- NOT all sent test commands are logged
	'							- NOT (probably) stepwise execution with logging defined at each step
	'							- NOT (probably) "misc" detailed schemes yet to be thought of
	'							- NOT all test output received back
	'							- All test results with 
	'							- NOT all output/lookfor comparisons on failure
	'				normal:		- 
	'				
	'				
	' console output:
	'		Complete text seen, same as logfile
	'		TODO: options for summary/normal/detailed console output

End Sub ' Private Sub RunTest(ByVal filename)

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
	'	NOTE: The resulting line is NOT right-trimmed, only left
	'	NOTE: Left-trimming was also removed, but not properly tested yet: 2018-06-15 V0.05-01
	'	TODO: Test whether removal of Left-trimming affected anything
	replace_args = argline
	f_name = "replace_args"
	f_error = "Error(" & f_name & "): "
	' say "replace_args called, argline="&argline		' *** DEBUG
	
	' argline  = trim(argline)
	if len(Trim(argline)) = 0 then Exit Function		' Empty line is OK, we just quit silently
	
	'remline = LTrim(argline)
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
			' say "ExecuteGlobal (""tmp_v=""&trim("&s&"))"		' *** DEBUG
			ExecuteGlobal "tmp_v="&trim(s)
			' tmp_v = Eval(s)
			If Err.Number <> 0 Then
				sayerr f_error & "Unable to substitute variable: """ & trim(s) & """"
				sayerr Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
				Exit Function
			End If
			On Error Goto 0

			argl2 = argl2 & tmp_v
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

End Function ' Function replace_args (argline)

sub sppa(s): say replace_args(s): end sub		' *** DEBUG

' Position of first non-space char in s
' or = len(s) +1 if none
Function Ltrimpos (s)
	LTrimPos = len (s) - len (Ltrim(s)) + 1
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
	preprocess_cmdline = ""
	f_name = "preprocess_cmdline"
	f_error = "Error(" & f_name & "): "
	' say "preprocess_cmdline called, cmdline="&cmdline		' *** DEBUG

	cmd = ""
	argline  = ""

	' Remove initial "." dot if present, but not spaces before of after dot
	' cmdline = LTrim (cmdline)
	ltpos = LTrimPos (cmdline) ' position of first non-blank
	'if Left (LTrim (cmdline), 1) = "." Then cmdline = LTrim (Mid (cmdline, 2))


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

	' say "Resulting command line=/"&cmline&"/"		' *** DEBUG

	preprocess_cmdline = cmd & " " & """"&argl2&""""

End Function ' Function preprocess_cmdline (cmdline)

 ' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
 Private Sub RunTestFile (ByVal filename)
	' RunTestFile - Similar to the "Import" routine, but some specific tricks apply:
	' ============================================================================
	' 
	' Preprocessing happens as follows for each line read from the file RunTestFile:
	'  0. "Normal lines": By default lines are executed as normal vbs code witout change
	'  1. "Dot substitution": Lines beginning with dot "." are converted as follows:
	'     Think of initial "." dot as: 
	'        Please run this command cmd ... for me, and btw I was too lazy to write quotes, so
	'        could you please add quotes around whatever comes after the command, and while you're
	'        at it, I'd appreciate if you would substitute expressions within curly braces too.
	'     This translates into the following steps:  
	'     a. Double quotes existing in the line are doubled, from " to ""
	'     b. Expressions enclosed in curly braces {} are evaluated and substituted
	'        example: "linecount = {linecount}" the value of linecount is substituted
	'        or "Date: {date} Time: {time}" will make those substitutions
	'     c. Curly braces within quotes are not substituted
	'     d. Quotes in expressions within curly braces are not doubled, in fact they are executed,
	'        as part of evaluating the expression.
	'     e. Then, initial "." removed, and first token separated, the remaining line in enclosed
	'        in quotes and the resulting line executed:
	'        eg. For the line: .say time={time} results in substitution of {time} with current time
	'     	 giving a line like "say ""time=15:12:36"""" which is then executed, giving
	'        the console output: time=15:12:36
	'     
	'     
	'     
	'  2. "right arrow:" Lines beginning with ">" or "_" are handled as follows:
	'     a. All double quotes existing in the line are doubled, from " to ""
	'     b. Then, initial ">" or "_" removed, the line is enclosed in ssend ("<line>")
	'        eg. The line: .setmem B7F8 "this is a string"
	'        becomes: ssend ("setmem B7F8 ""this is a string""")
	'     
	'! 
	'! 		a. 
	'!		2. 
	'!		
	'!		
	'!		and then executed.
	'!		ssend will then pass the string in as a command line to the slave app
	'!		with the doubled quotes converted back to single ones again


	Dim exitOnError, fso, sh, file, code, dir, executeNow, nonExeCount, fLine, TestLine, qpos, ppos

	filename = "  " & filename & "   "
	getopts ":E", filename, opts_found

	exitOnError = Not find_opt("E", opts_found)

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")
	executeNow	= True
	nonExeCount = 0
	fLine	 	= ""
	TestLine 	= ""
	code     	= ""


	' Find filename and open it
	' --------------------------------------------------------------------------------
	saydbg "@all Initial filename="&filename
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
	
	Set file = fso.OpenTextFile(filename, 1, False)
	flineno = 0

	' Main loop: Read from file, line by line
	' ----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+----+
	Do While Not file.AtEndOfStream 
		fLine 		= file.ReadLine		' An untrimmed fLine may be needed below in some cases
		flineno = flineno + 1
		TestLine 	= LTrim(fLine)
		
		' Handle multi-line code block
		' ----+----+----+----+----+----+----+----+----+----L----+----+----+----+----+----+----+----+----+----C
		' ----.----o----.----o----.----o----.----o----.----L----+----o----+----o----+----o----+----o----+----|
		' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
		' ====+====1====+====2====+====3====+====4====+====5====+====6====+====7====+====8====+====9====+====0
		' This is a multi-line block of code, enclosed by start and end tokens
		' that MUST be read in from file FIRST, and THEN executed as ONE CHUNK
		' The "smiley bird" symbols "<:" and ":>" respectively are the start and end tokens
		If Left (TestLine, 2) = "<:" Then		' Left <: smiley-bird block indicator 
			right_smiley_found = False
			' TestLine = Mid (TestLine, 3)
			TestLine = Mid (LTrim(fline), 3)
			Do
				' Handle case where end token ":>" is on same line
				' Uncertain, TestCase: what if fLine has end token followed by spaces(?)
				' sayerr "smiley-bird block line: " & TestLine
				If Right(Testline, 2) = ":>" Then
					Testline = Left(TestLine, len (TestLine) - 2 )
					right_smiley_found = True
				End If

				' Preprocess
				'Testline = Trim (TestLine)
				If Mid(Testline, LTrimPos(Testline), 1) = "." Then
					' sayq "preprocessing:"&TestLine  ' **** DEBUG
					Testline = preprocess_cmdline (TestLine)
				End If

				'sayq "Testline="&Testline  ' **** DEBUG

				If len (code) > 0 Then code = code & VBCrLf 
				code = code & TestLine

				' Exit the loop if finished
				If right_smiley_found Or file.AtEndOfStream Then Exit Do

				' Get next line
				TestLine = file.ReadLine
				flineno = flineno + 1
			Loop
			' sayerr "smiley-bird block found:" & vbcrlf & "----------" & vbcrlf & code & vbcrlf & "----------"
			TestLine = ""
		End If	' If Left (TestLine, 2) = "<:" Then		' Left <: smiley-bird block indicator 

		' Handle "here-block" (aka "here document")
		' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
		' The "here-block" is a multi-line block of text, enclosed in "<+" and "+>" respectively
		ppos = InStr (TestLine, "<+")
		If ppos > 0 Then		' Left <+ here-block indicator 
			'sayq "TestLine="&TestLine
			'sayq "ppos="&ppos
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

			'sayq "qcount="&qcount
		
			If qcount mod 2 = 0 Then
				'say "Even number of quotes detected, or none: Thus, a VALID here-block start marker was found"

				code    = Left(TestLine, ppos - 1)
				remline = Mid (TestLine, ppos + 2)

				if len(remline) > 0 then
					code = code & """" & replace_args (remline) & """ & VBCrLf"
				Else
					code = code & """"""
				End If 
				'sayerr "code="&code


				Do
					' fLine = code: sayerr "here block line: " & fLine
					fLine = file.ReadLine
					flineno = flineno + 1

					If Right(fLine, 2) = "+>" Or file.AtEndOfStream Then
						fLine = Left(fLine, len (fLine) - 2 )
						code = code & " & """ & replace_args (fLine) & """"
						Exit Do
					Else 
						code = code & " & """ & replace_args (fLine) & """ & VBCrLf"
					End If
				Loop
				'sayerr "here block found:" & vbcrlf & "----------" & vbcrlf & code & vbcrlf & "----------"


				TestLine = ""
			End If
		End If	' If ppos > 0 Then		' Left <+ here block indicator 

		' Handle "Underscore-continued" blocks
		' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
		' multi-line block of code, signalled by space+underscore " _" at the end of each line,
		' or just an underscore in the case of empty lines
		' Note: The input lines are concatenated into one long line, space-separated, without any linefeeds
		' TODO: Analyse possible conflict between Underscore-continued lines and smiley-bird blocks
		if len (code) = 0 Then code = TestLine
		Do While Right(code, 2) = " _" Or Right(code, 3) = VBCrLf & "_"  Or code = "_"
			code = RTrim(Left(code, Len(code)-1)) & " " & Trim(file.ReadLine)
			flineno = flineno + 1
		Loop

		' Handle lines to be piped to the stdin of a slave process
		' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
		' Lines beginning with a slave marker "_" or ">" are sent to stdin of slave app
		' The lines are preprocessed with quote-doubling, {} curly-brace substitution and quote-enclosed
		' TODO: if the slave itself is another lookfor instance, it should refuse these lines
		'		at this point. Slave-of-slave or multi-slave scenarios don't seem worth while for now
		' TODO: check slave-sending of multiline code, underscore-continued or <:multi-line:> code blocks
		If Left (code, 1) = "_" Or Left (code, 1) = ">" Then
			code = Mid(code,2)
			code = replace_args (code)
			'code = Replace(code,"""","""""")
			
			'sayerr "Test output: ssend(" & code & ")"		' *** DEBUG
			'code = "ssend(" & code & ")"
			
			code = "ssend(""" & code & """)"
		End If
		

		' "normal" Lines are executed locally ie. right here

		' SPOX (Single Point of Execution): Execute the line, or block of lines that were read in
		' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
		If executeNow Then
			' Boolean executenow is for handling  multiline input, currently always True ie. NOT USED
			nonExeCount = 0
			MyErr.Reset
			RunVbshLine(code)

			If MyErr.Number <> 0 Then
				sayerr "(RunTestFile): Execution Error in file: " & shortfilename & ", Line no: " & flineNo
				If exitOnError Then Exit Sub
			End If

			code = ""
		ElseIf Not executeNow Then		' NOTE: Currently NOT used 
			nonExeCount = nonExeCount+1

			' ====== DEAD CODE ======
		Else ' This Else is never reached, left in to keep code below as reference
			On Error Resume Next
				Err.Clear
				ExecuteGlobal(TestLine)
				If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
			' ====== END of DEAD CODE ======
		End If

	'		code = code & TestLine & vbCrLf
	' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
	Loop ' Do While Not file.AtEndOfStream 

	'	code = code & "  "

	file.Close()

	'! say "processed code from file " & filename & ":"
	'! say code


	'	On Error Resume Next
	'		Err.Clear
	'		ExecuteGlobal(code)
	'		If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
	'	On Error Goto 0

	Set fso = Nothing
	Set sh  = Nothing

  ' ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----O
 End Sub '! Private Sub RunTestFile(ByVal filename)

 sub rtf(s): RunTestFile s: end sub		' *** DEBUG

 Function sans_prompt (s, prompt)
	sans_prompt = Replace (s, vbCrLf & prompt, "") '! Not sure if we should remove only one here (?)
	'!sans_prompt = Replace (s, prompt) '! Not sure if we should remove these (?)
 End Function '! Function sans_prompt (s, prompt)

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
 ' Old stuff kept around in here for possible reference or, more likely, to be thrown out soon
 ' ====================================================================================================
 ' ====================================================================================================
Private Sub old_RunTestFile(ByVal filename)
	'! RunTestFile - Same as "Import" above, but some preprocessing is done:
	'! ============================================================================
	'! By default lines are not changed. They are executed as normal vbs code
	'! Preprocessing happens as follows:
	'! Lines beginning with "." or ">" are converted as follows:
	'! 		1. All double quotes existing in the line are doubled, from " to ""
	'!		2. Then, initial dot removed, the line is enclosed in ssend ("<line>")
	'!		eg. The line: .setmem B7F8 "this is a string"
	'!		is converted to: ssend ("setmem B7F8 ""this is a string""")
	'!		and then executed.
	'!		ssend will then pass the string in as a command line to the slave app
	'!		with the doubled quotes converted back to single ones again
	Dim fso, sh, file, code, dir

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	if InStr(filename, ".") = 0 then
	   filename = filename & ".vbst"
	end if
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
	
	TestLine = ""
	code     = ""

	Set file = fso.OpenTextFile(filename, 1, False)


	Do While Not file.AtEndOfStream 
		fLine 		= file.ReadLine		' An untrimmed fLine may be needed below in some cases
		TestLine 	= Trim(fLine)
		
		' Handle "smiley bird"-enclosed code
		' TODO: preprocess, so you can use vbsx code in a smiley bird block too
		' vbsx code: .cmd <text to be preprocessed and enclosed in quotes>
		' TODO: "here-docs in vbsx code"
		If Left (TestLine, 2) = "<:" Then
			Testline = Mid (TestLine, 3)
			Do
				' Handle case where right-smiley-bird is on same line
				' Uncertain, TestCase: what if fLine has right-smiley followed by spaces(?)
				If Right(Testline, 2) = ":>" Then
					Testline = Left(TestLine, len (TestLine) - 2 )
					Exit Do
				End If
				TestLine = TestLine & VBCrLf & Trim(file.ReadLine)
			Loop
		End If

		Do While Right(TestLine, 2) = " _" Or TestLine = "_"
			TestLine = RTrim(Left(TestLine, Len(TestLine)-1)) & " " & Trim(file.ReadLine)
		Loop

		' Lines beginning with a slave marker "_" or ">" are sent to stdin of slave app
		' TODO: if the slave itself is another lookfor instance, it should refuse these lines
		If InStr (TestLine, "_") = 1 or InStr (TestLine, ">") = 1 Then
			TestLine = Mid(TestLine,2)
			TestLine = replace_args (TestLine)
			'TestLine = Replace(TestLine,"""","""""")
			
			'sayerr "Test output: ssend(" & TestLine & ")"		' *** DEBUG
			'TestLine = "ssend(" & TestLine & ")"
			
			TestLine = "ssend(""" & TestLine & """)"
		End If

		' "normal" Lines are executed locally ie. right here

		If true then 
			RunVbshLine(TestLine)
		else
			On Error Resume Next
				Err.Clear
				ExecuteGlobal(TestLine)
				If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
		End If

	'		code = code & TestLine & vbCrLf
	Loop ' Do While Not file.AtEndOfStream 

	'	code = code & "  "

	file.Close()

	'! say "processed code from file " & filename & ":"
	'! say code


	'	On Error Resume Next
	'		Err.Clear
	'		ExecuteGlobal(code)
	'		If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
	'	On Error Goto 0

	Set fso = Nothing
	Set sh  = Nothing

End Sub '! Private Sub old_RunTestFile(ByVal filename)

