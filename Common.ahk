; =========
; Functions
; =========

/*
  Kills other AutoHotKey scripts
*/
AHKPanic(Kill=0, Pause=0, Suspend=0, SelfToo=0) 
{
  DetectHiddenWindows, On
  WinGet, IDList ,List, ahk_class AutoHotkey
  Loop %IDList% 
  {
    ID:=IDList%A_Index%
    WinGetTitle, ATitle, ahk_id %ID%
    IfNotInString, ATitle, %A_ScriptFullPath% 
    {
      If Suspend
        PostMessage, 0x111, 65305,,, ahk_id %ID% ; Suspend
      If Pause
        PostMessage, 0x111, 65306,,, ahk_id %ID% ; Pause
      If Kill
        WinClose, ahk_id %ID% ;kill
    }
  }
  If SelfToo 
  {
    If Suspend
      Suspend, Toggle  ; Suspend
    If Pause
      Pause, Toggle, 1 ; Pause
    If Kill
      ExitApp
  }
}

/*
  Wait for a window to become active. Focuses it if it's there already.
*/
WaitForWin(WindowTitle) 
{
  WinWait, %WindowTitle%, , WinWaitActiveDelay 
  IfWinNotActive, %WindowTitle%, , WinActivate, %WindowTitle%,
  WinWaitActive, %WindowTitle%,
}

/*
  Adds a line group in the specified item # in the list
*/
AddLineGroupToRegionalDialPlan(LineGroupName, EntryPosition)
{
  ; Select item in dial plan list
  Send %EntryPosition%

  ; Edit it
  ControlClick, Button3
  WaitForWin("Regional Dial Plan - Edit Pattern")
  Sleep WindowActiveDelay

  ; Click on Add Group
  ControlClick, Button8
  WaitForWin("Dial Group - Add Entry")
  Sleep WindowActiveDelay

  ; Select Loopback
  ControlFocus, SysListView321, Dial Group - Add Entry
  SetKeyDelay InputKeyDelay
  Send %LineGroupName%
  SetKeyDelay KeyDelay

  ; Click on OK
  ControlClick, Button2

  ; Wait For Regional Dial Plan - Edit Pattern to come back
  WaitForWin("Regional Dial Plan - Edit Pattern")
  Sleep WindowActiveDelay

  ; Click on OK
  ControlClick, Button17

  ; Wait For Regional Dial Plan to come back
  WaitForWin("Regional Dial Plan")
  Sleep WindowActiveDelay
}