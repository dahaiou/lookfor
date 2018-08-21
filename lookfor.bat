@echo off
rem cf: vbsh at these url's
rem https://stackoverflow.com/questions/15087377/how-can-i-start-an-interactive-console-for-vbs
rem http://www.planetcobalt.net/sdb/vbsh.shtml
rem https://kryogenix.org/days/2004/04/01/interactivevbscript/

rem @cscript.exe //NoLogo lookfor.vbs

set batpath=%~dp0
set batpath=%batpath:~0,-1%

rem echo batpath is: %batpath%

@cscript.exe 2>&1 /NoLogo %batpath%\lookfor.vbs %*
