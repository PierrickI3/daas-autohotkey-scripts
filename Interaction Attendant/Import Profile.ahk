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
KeyDelay 			 = 500
InputKeyDelay 		 = 50
WinWaitActiveDelay 	 = 300
WindowActiveDelay    = 200
CloseAttendantOnExit = false

SetKeyDelay KeyDelay

; Kill Other Scripts
AHKPanic(1, 0, 0, 0)

; ===========================
; Start Interaction Attendant
; ===========================
; If not already started, start it
Process, Exist, InteractionAttendantU.exe
If (ErrorLevel == 0) ; 
{
	CloseAttendantOnExit = true
	Run "C:\I3\IC\Server\InteractionAttendantU.exe"
}
WaitForWin("Interaction Attendant")
Sleep, 2000

; Open File/Import File menu option
SetKeyDelay, InputKeyDelay
Send {AltDown}F{AltUp}
Sleep, 500
Send, I
Sleep, 500
SetKeyDelay, KeyDelay
WaitForWin("Open")
Sleep, 500

; Enter package filename and press Enter
SetKeyDelay, InputKeyDelay
Send, %1%
Send, {Enter}
SetKeyDelay, KeyDelay
WaitForWin("Import Configuration")

; Select server
SetKeyDelay, InputKeyDelay
Send, {Tab}{Tab}{Down}{Down}{Enter}
SetKeyDelay, KeyDelay
WaitForWin("Attendant")

; Skip warning about being already connected
Send, {Enter}
WaitForWin("Attendant")

; Close "You have successfully imported..." dialog box
Send, {Enter}
WaitForWin("Interaction Attendant")
Sleep, 500

; Publish new configuration
SetKeyDelay, InputKeyDelay
Send, {AltDown}F{AltUp}
Sleep, 500
Send, P
Sleep, 500
WaitForWin("Attendant")
Sleep, 200
Send, {Enter}
SetKeyDelay, KeyDelay
WaitForWin("Interaction Attendant")
Sleep, 5000 ; Wait for publish to complete (?? Better way ??)

if (CloseAttendantOnExit) {
	; Close Interaction Attendant
	SetKeyDelay, InputKeyDelay
	Send, {AltDown}F{AltUp}
	Sleep, 500
	Send, X
}
