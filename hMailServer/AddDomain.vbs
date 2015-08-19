' Global variables
Public obApp
Public obAdminAccount
Public obNewDomain

Public Const user = "Administrator"
Public Const password = ""

' Check if a domain name was passed
Set oArgs = WScript.Arguments
If oArgs.count <> 1 Then
   WScript.Quit
End If

' Authenticate
Set obApp = CreateObject("hMailServer.Application")
Set obAdminAccount = obApp.Authenticate(user, password)

' Add new domain
Set obNewDomain = obApp.Domains.Add()
obNewDomain.Name = oArgs(0)
obNewDomain.Active = True
obNewDomain.Save()

obApp.Domains.Refresh()