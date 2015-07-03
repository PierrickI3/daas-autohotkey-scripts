#SingleInstance force
#NoEnv  						; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%  	; Ensures a consistent starting directory.
StringCaseSense On
AutoTrim OFF

; ========
; Includes
; ========
#Include ..\Common.ahk

; =========
; Variables
; =========
KeyDelay 			= 500
InputKeyDelay 		= 50
WinWaitActiveDelay 	= 300
LineName 			= Loopback
LineGroupName 		= Loopback

SetKeyDelay KeyDelay

; Kill Other Scripts
AHKPanic(1, 0, 0, 0)

; ===============================
; Start Interaction Administrator
; ===============================
; Close Existing instance of Interaction Administrator
Process, Exist, IAShellU.exe
If (ErrorLevel != 0) ; 
{
  WinClose Interaction Administrator
}

; Open Interaction Administrator
Run "C:\I3\IC\Server\IAShellU.exe"
WaitForWin("Interaction Administrator")
Sleep, 2000

; ============
; Create Lines
; ============
; Go to lines
SetKeyDelay InputKeyDelay
Send Lines
SetKeyDelay KeyDelay

; Create a new line
Sleep 200
Send ^n
WaitForWin("Entry Name")

; Enter Line name
SetKeyDelay InputKeyDelay 
Send %LineName%{enter}
SetKeyDelay KeyDelay

; Wait for Line Configuration dialog
WaitForWin("Line Configuration")

; Go to Identity (Out)
Send {down}{down}
Sleep 300

; Open Line Value 1
ControlClick, Button22

; Wait for Configure Line Value dialog
WaitForWin("Configure Line Value")

; Check Anonymous
Send {space}{enter}

; Wait for Line Configuration dialog to come back
WaitForWin("Line Configuration")

; Get back to menu
Send {SHIFT DOWN}{tab}{tab}{tab}{SHIFT UP}

; Go to Proxy
Send {down}{down}{down}{down}{down}
Sleep 200

; Click on Add
ControlClick, Button35

; Wait for SIP Address dialog
WaitForWin("SIP Address")

; Set to local IP address of "Ethernet 2" adapter
SetKeyDelay InputKeyDelay 
Send %A_IPAddress1%{enter}
SetKeyDelay KeyDelay 

; Wait for Line Configuration dialog to come back
WaitForWin("Line Configuration")

; We're done, press ok
ControlClick, Button67

; =================
; Create Line Group
; =================
; Go to Line Groups
SetKeyDelay InputKeyDelay
Send Line Groups
SetKeyDelay KeyDelay

; Create a new line group
Sleep 200
Send ^n

; Wait for Entry Name dialog
WaitForWin("Entry Name")

; Enter Line Group Name
SetKeyDelay InputKeyDelay 
Send %LineGroupName%{enter}
SetKeyDelay KeyDelay 

; Select for reporting and Dial Group
Sleep 100
Send {tab}{space}{tab}{space}

; Go to Members tab
Send {CTRL DOWN}{tab}{CTRL UP}

; Select "Loopback" entry
Sleep 100
Control, ChooseString, %LineName%, ListBox1, Line Group Configuration

; Close Line Group
Send {tab}{enter}
WaitForWin("Interaction Administrator")

; ================
; Add to Dial Plan
; ================
; Go to Phone Numbers
SetKeyDelay InputKeyDelay
Send Phone Numbers
SetKeyDelay KeyDelay

; Edit Configuration
Send {tab}{down}{enter}
WaitForWin("Phone Number Configuration")

; Open Dial Plan
Send {space}
WaitForWin("Regional Dial Plan")

; ===========================
; Start Line Group Assignment
; ===========================
; Select sip:Z
Send {tab}1 ; sip:Z is in first position. Find a better way?

; Edit it
ControlClick, Button3
WaitForWin("Regional Dial Plan - Edit Pattern")

; Click on Add Group
ControlClick, Button8
WaitForWin("Dial Group - Add Entry")

; Select Loopback
ControlFocus, SysListView321, Dial Group - Add Entry
SetKeyDelay InputKeyDelay
Send %LineName%
SetKeyDelay KeyDelay

; Click on OK
ControlClick, Button2

; Wait For Regional Dial Plan - Edit Pattern to come back
WaitForWin("Regional Dial Plan - Edit Pattern")

; =========================
; End Line Group Assignment
; =========================

; Click on OK
ControlClick, Button17

; Wait For Regional Dial Plan to come back
WaitForWin("Regional Dial Plan")

;Exit (click on OK)
ControlClick, Button11

; Wait For Phone Number Configuration to come back
WaitForWin("Phone Number Configuration")

;Exit (click on OK)
ControlClick, Button1