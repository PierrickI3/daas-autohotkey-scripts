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
  ControlFocus, SysListView321, Regional Dial Plan
  
  ; Select item in dial plan list
  SetKeyDelay 50
  Send %EntryPosition%
  SetKeyDelay 500
  Sleep 200

  ; Edit it
  ControlClick, Button3
  WaitForWin("Regional Dial Plan - Edit Pattern")
  Sleep 200

  ; Click on Add Group
  ControlClick, Button8
  WaitForWin("Dial Group - Add Entry")
  Sleep 200

  ; Select Loopback
  ControlFocus, SysListView321
  SetKeyDelay 10
  Send %LineGroupName%
  SetKeyDelay 500

  ; Click on OK
  Sleep 200
  ControlClick, Button2

  ; Wait For Regional Dial Plan - Edit Pattern to come back
  WaitForWin("Regional Dial Plan - Edit Pattern")
  Sleep 200

  ; Click on OK
  ControlClick, Button17

  ; Wait For Regional Dial Plan to come back
  WaitForWin("Regional Dial Plan")
  Sleep 200
}

/*
  Removes a line group in the specified item # in the list
*/
RemoveLineGroupToRegionalDialPlan(LineGroupName, EntryPosition)
{
  ControlFocus, SysListView321, Regional Dial Plan
  
  ; Select item in dial plan list
  SetKeyDelay 50
  Send %EntryPosition%
  SetKeyDelay 500
  Sleep 200

  ; Edit it
  ControlFocus, Button3, Regional Dial Plan
  ControlClick, Button3, Regional Dial Plan
  WaitForWin("Regional Dial Plan - Edit Pattern")
  Sleep 200

  ; Select Loopback
  ControlFocus, SysListView321
  Sleep 200
  SetKeyDelay 50
  Send %LineGroupName%
  SetKeyDelay 500

  ; Click on Remove
  Sleep 200
  ControlClick, Button10
  Sleep 200

  ; Click on OK
  ControlClick, Button17

  ; Wait For Regional Dial Plan to come back
  WaitForWin("Regional Dial Plan")
  Sleep 200
}
