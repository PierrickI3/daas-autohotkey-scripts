#SingleInstance force
#NoEnv  			        ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
StringCaseSense On
AutoTrim OFF

; ========
; Includes
; ========
#Include ..\Common.ahk

; Kill Other Scripts
AHKPanic(1, 0, 0, 0)

; Variables
KeyDelay 		    = 500
InputKeyDelay 		= 50
WinWaitActiveDelay 	= 300
WindowActiveDelay   = 200
LineName 		    = Loopback
LineGroupName 		= Loopback

SetKeyDelay KeyDelay

; ====================================================
; Close Existing instance of Interaction Administrator
; ====================================================
Process, Exist, IAShellU.exe
If (ErrorLevel != 0) ; 
{
  WinClose Interaction Administrator
}

Run "C:\I3\IC\Server\IAShellU.exe"
WaitForWin("Interaction Administrator")
Sleep, 500

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

; Select sip:Z
Send {tab}1 ; sip:Z is in first position

; Edit it
ControlClick, Button3
WaitForWin("Regional Dial Plan - Edit Pattern")
Sleep WindowActiveDelay

; Go to the list view
Send {SHIFT DOWN}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{SHIFT UP}

; Select Loopback
ControlFocus, SysListView321, Dial Group - Add Entry
SetKeyDelay InputKeyDelay
Send %LineName%
SetKeyDelay KeyDelay

; Click on Remove
Sleep 100
ControlClick, Button10
Sleep 100

; Click on OK
ControlClick, Button17

; ======================
; End Line Group Removal
; ======================

WaitForWin("Regional Dial Plan")
Sleep WindowActiveDelay

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
