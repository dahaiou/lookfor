@echo off
rem cf: vbsh at these url's
rem https://stackoverflow.com/questions/15087377/how-can-i-start-an-interactive-console-for-vbs
rem http://www.planetcobalt.net/sdb/vbsh.shtml
rem https://kryogenix.org/days/2004/04/01/interactivevbscript/

rem This is cmd_lookfor.bat. Call it from Windows start menus and such,
rem to open a new cmd console window and run lookfor.vbs in it
rem This console stays open even if you exit the lookfor.vbs script,
rem and you can simply go "lookfor" to run it again

rem From an already open cmd console, call lookfor.bat instead

cmd.exe /K @cscript.exe //NoLogo lookfor.vbs %*
