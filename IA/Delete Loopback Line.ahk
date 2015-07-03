#SingleInstance force
#NoEnv  			        ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
StringCaseSense On
AutoTrim OFF

; ========
; Includes
; ========
#Include ..\Common.ahk

; =========
; Variables
; =========
KeyDelay 		    = 500
InputKeyDelay 		= 50
WinWaitActiveDelay 	= 300
WindowActiveDelay   = 200
LineName 		    = Loopback
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

Run "C:\I3\IC\Server\IAShellU.exe"
WaitForWin("Interaction Administrator")
Sleep, 2000
Send {Home}

; =====================
; Remove from Dial Plan
; =====================
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

; =======================================
; Start Line Group Removal from Dial Plan
; =======================================
Send {tab}

; Only works with the US Dial Plan with one local exchange!

; Remove from sip:Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From sips:Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From tel:Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 3)
; Remove From ?@Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 4)
; Remove From 00
RemoveLineGroupToRegionalDialPlan(LineGroupName, 6)
; Remove From 911
RemoveLineGroupToRegionalDialPlan(LineGroupName, 7)
; Remove From 411
RemoveLineGroupToRegionalDialPlan(LineGroupName, 8)
; Remove From 1411
RemoveLineGroupToRegionalDialPlan(LineGroupName, 9)
; Remove From 5551212
RemoveLineGroupToRegionalDialPlan(LineGroupName, 10)
; Remove From Nxx5551212
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From 1Nxx5551212
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From +1Nxx5551212
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From NXXXXXXXXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From $$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From 1$$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From +1$$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From $$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From 1$$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 1)
; Remove From +1$$$MxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 20)
; Remove From 317XXXXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From 1317XXXXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From +1317XXXXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From NxxNxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From 1NxxNxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From +1NxxNxxXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From NXXXXXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From $$$XXXXZ
RemoveLineGroupToRegionalDialPlan(LineGroupName, 2)
; Remove From 011+Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 36)
; Remove From 011Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 3)
; Remove From +Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 3)
; Remove From Z
RemoveLineGroupToRegionalDialPlan(LineGroupName, 3)

; ======================
; End Line Group Removal
; ======================

; Click on OK
ControlClick, Button11

WaitForWin("Phone Number Configuration")
Sleep WindowActiveDelay

; Click on OK
ControlClick, Button1

; =================
; Remove Line Group
; =================
; Go to Line Groups
Send {SHIFT DOWN}{tab}{SHIFT UP}
SetKeyDelay InputKeyDelay
Send Line Groups
SetKeyDelay KeyDelay
Sleep 100

; Select Loopback line
Send {tab}
SetKeyDelay InputKeyDelay
Send %LineGroupName%
SetKeyDelay KeyDelay

; Delete it!
Send {delete}
Sleep 100
WaitForWin("Interaction Administrator")
Sleep WindowActiveDelay
ControlClick, Button1

; ===========
; Remove Line
; ===========
; Go to Lines
Send {SHIFT DOWN}{tab}{SHIFT UP}
SetKeyDelay InputKeyDelay
Send Lines
SetKeyDelay KeyDelay
Sleep 100

; Select Loopback line
Send {tab}
SetKeyDelay InputKeyDelay
Send %LineName%
SetKeyDelay KeyDelay
Sleep 100

; Delete it!
Send {delete}
Sleep 100

; Confirm
WaitForWin("Interaction Administrator")
Sleep WindowActiveDelay
ControlClick, Button1
