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
WindowActiveDelay   = 200
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
Sleep WindowActiveDelay

; Enter Line name
SetKeyDelay InputKeyDelay 
Send %LineName%{enter}
SetKeyDelay KeyDelay

; Wait for Line Configuration dialog
WaitForWin("Line Configuration")
Sleep WindowActiveDelay

; Go to Identity (Out)
Send {down}{down}
Sleep WindowActiveDelay

; Open Line Value 1
ControlClick, Button22

; Wait for Configure Line Value dialog
WaitForWin("Configure Line Value")
Sleep WindowActiveDelay

; Check Anonymous
Send {space}{enter}

; Wait for Line Configuration dialog to come back
WaitForWin("Line Configuration")
Sleep WindowActiveDelay

; Get back to menu
Send {SHIFT DOWN}{tab}{tab}{tab}{SHIFT UP}

; Go to Proxy
Send {down}{down}{down}{down}{down}
Sleep WindowActiveDelay

; Click on Add
ControlClick, Button35

; Wait for SIP Address dialog
WaitForWin("SIP Address")
Sleep WindowActiveDelay

; Set to local IP address of "Ethernet 2" adapter
SetKeyDelay InputKeyDelay 
Send %A_IPAddress1%{enter}
SetKeyDelay KeyDelay 

; Wait for Line Configuration dialog to come back
WaitForWin("Line Configuration")
Sleep WindowActiveDelay

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
Sleep WindowActiveDelay
Send ^n

; Wait for Entry Name dialog
WaitForWin("Entry Name")
Sleep WindowActiveDelay

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
Sleep WindowActiveDelay

; ================
; Add to Dial Plan
; ================
; Go to Phone Numbers
SetKeyDelay InputKeyDelay
Send Phone Numbers
SetKeyDelay KeyDelay

; Edit Configuration
Sleep 100
Send {tab}{down}{enter}
WaitForWin("Phone Number Configuration")
Sleep WindowActiveDelay

; Open Dial Plan
Send {space}
WaitForWin("Regional Dial Plan")
Sleep WindowActiveDelay

; ===========================
; Start Line Group Assignment
; ===========================
Send {tab}

; Only works with the US Dial Plan with one local exchange!

; Add to sip:Z
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to sips:Z
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to tel:Z
AddLineGroupToRegionalDialPlan(LineGroupName, 3)
; Add to ?@Z
AddLineGroupToRegionalDialPlan(LineGroupName, 4)
; Add to 00
AddLineGroupToRegionalDialPlan(LineGroupName, 6)
; Add to 911
AddLineGroupToRegionalDialPlan(LineGroupName, 7)
; Add to 411
AddLineGroupToRegionalDialPlan(LineGroupName, 8)
; Add to 1411
AddLineGroupToRegionalDialPlan(LineGroupName, 9)
; Add to 5551212
AddLineGroupToRegionalDialPlan(LineGroupName, 10)
; Add to Nxx5551212
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to 1Nxx5551212
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to +1Nxx5551212
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to NXXXXXXXXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to $$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to 1$$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to +1$$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to $$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to 1$$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 1)
; Add to +1$$$MxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 20)
; Add to 317XXXXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to 1317XXXXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to +1317XXXXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to NxxNxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to 1NxxNxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to +1NxxNxxXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to NXXXXXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to $$$XXXXZ
AddLineGroupToRegionalDialPlan(LineGroupName, 2)
; Add to 011+Z
AddLineGroupToRegionalDialPlan(LineGroupName, 36)
; Add to 011Z
AddLineGroupToRegionalDialPlan(LineGroupName, 3)
; Add to +Z
AddLineGroupToRegionalDialPlan(LineGroupName, 3)
; Add to Z
AddLineGroupToRegionalDialPlan(LineGroupName, 3)

; =========================
; End Line Group Assignment
; =========================

;Exit (click on OK)
ControlClick, Button11

; Wait For Phone Number Configuration to come back
WaitForWin("Phone Number Configuration")
Sleep WindowActiveDelay

;Exit (click on OK)
ControlClick, Button1