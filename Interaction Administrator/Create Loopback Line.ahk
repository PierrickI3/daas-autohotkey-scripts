#SingleInstance force
#NoEnv  						; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  				; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  	; Ensures a consistent starting directory.
StringCaseSense On
AutoTrim OFF

; ========
; Includes
; ========
#Include ..\Common.ahk

; Kill Other Scripts
AHKPanic(1, 0, 0, 0)

; Variables
KeyDelay 			= 500
InputKeyDelay 		= 50
WinWaitActiveDelay 	= 300
LineName 			= Loopback
LineGroupName 		= Loopback

SetKeyDelay KeyDelay 

; Open Interaction Administrator
Run "C:\I3\IC\Server\IAShellU.exe"
WaitForWindow("Interaction Administrator")
Sleep, 2000

; ============
; Create Lines
; ============
; Go to lines
Send {down}{down}{down}{down}
; Create a new line
Send ^n
WaitForWindow("Entry Name")
; Enter Line name
SetKeyDelay InputKeyDelay 
Send %LineName%{enter}
SetKeyDelay KeyDelay 
; Wait for Line Configuration dialog
WaitForWindow("Line Configuration")

; Go to Identity (Out)
Send {down}{down}
; Open Line Value 1
ControlClick, Button22
; Wait for Configure Line Value dialog
WaitForWindow("Configure Line Value")
; Check Anonymous
Send {space}{enter}
; Wait for Line Configuration dialog to come back
WaitForWindow("Line Configuration")
; Get back to menu
Send {SHIFT DOWN}{tab}{tab}{tab}{SHIFT UP}

; Go to Proxy
Send {down}{down}{down}{down}{down}
; Click on Add
ControlClick, Button35
; Wait for SIP Address dialog
WaitForWindow("SIP Address")
; Set to local IP address of "Ethernet 2" adapter
SetKeyDelay InputKeyDelay 
Send %A_IPAddress1%{enter}
SetKeyDelay KeyDelay 
; Wait for Line Configuration dialog to come back
WaitForWindow("Line Configuration")
; We're done, press ok
ControlClick, Button67

; =================
; Create Line Group
; =================
Send {down}^n
; Wait for Entry Name dialog
WaitForWindow("Entry Name")
; Enter Line Group Name
SetKeyDelay InputKeyDelay 
Send %LineGroupName%{enter}
SetKeyDelay KeyDelay 
; Select for reporting and Dial Group
Send {tab}{space}{tab}{space}
; Go to Members tab
Send {CTRL DOWN}{tab}{CTRL UP}
; Select "Loopback" entry
Control, ChooseString, %LineName%, ListBox1, Line Group Configuration
; Close Line Group
Send {tab}{enter}
WaitForWindow("Interaction Administrator")

; ================
; Add to Dial Plan
; ================
; Click on Phone Numbers
Click, 259, 192
ControlClick, x115 y376, Interaction Administrator
Send {tab}{down}{enter}
WaitForWindow("Phone Number Configuration")

; Open Dial Plan
Send {space}
WaitForWindow("Regional Dial Plan")

; ===========================
; Start Line Group Assignment
; ===========================

; Select sip:Z
Send {tab}{down} ;TODO Find a better way? By string?

; Edit it
ControlClick, Button3
WaitForWindow("Regional Dial Plan - Edit Pattern")

; Click on Add Group
ControlClick, Button8
WaitForWindow("Dial Group - Add Entry")

; Select Loopback
ControlFocus, SysListView321, Dial Group - Add Entry
SetKeyDelay InputKeyDelay
Send %LineName%
SetKeyDelay KeyDelay
ControlClick, Button2

; Wait For Regional Dial Plan - Edit Pattern to come back
WaitForWindow("Regional Dial Plan - Edit Pattern")

; =========================
; End Line Group Assignment
; =========================

ControlClick, Button17

; Wait For Regional Dial Plan to come back
WaitForWindow("Regional Dial Plan")
;Exit (click on OK)
ControlClick, Button11

; Wait For Phone Number Configuration to come back
WaitForWindow("Phone Number Configuration")
;Exit (click on OK)
ControlClick, Button1
