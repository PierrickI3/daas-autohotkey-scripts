Set oArgs = WScript.Arguments 
If oArgs.count <> 1 Then 
   WScript.Quit 
End If 

Dim obBaseApp 
Set obBaseApp = CreateObject("hMailServer.Application") 

Dim obNewDomain 
Set obNewDomain = obBaseApp.Domains.Add() 

obNewDomain.Name = oArgs(0) 
obNewDomain.Active = True 
obNewDomain.Save() 

obBaseApp.Domains.Refresh()