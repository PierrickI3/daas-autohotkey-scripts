@echo off

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%

REM "Create Loopback Line" shortcut
REM -------------------------------
echo sLinkFile = "%USERPROFILE%\Desktop\Create Loopback Line.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "C:\Users\vagrant\Desktop\Scripts\AutoHotKey\IA\Create Loopback Line.ahk" >> %SCRIPT%
echo oLink.WorkingDirectory = "C:\Users\vagrant\Desktop\Scripts\AutoHotKey\IA" >> %SCRIPT%
echo oLink.Description = "Creates a loopback line in Interaction Administrator" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%

cscript /nologo %SCRIPT%
del %SCRIPT%
