Option Explicit

Dim obBaseApp
Dim objFSO
Dim objTextFile
Dim strNewAlias,i

Const ForReading = 1
 
Set obBaseApp = CreateObject("hMailServer.Application") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("Users.csv", ForReading)

Do While objTextFile.AtEndOfStream <> True
    strNewAlias = split(objTextFile.Readline, ",")

    Select Case strNewAlias(0)
       Case "User"
          AddUser strNewAlias(1), strNewAlias(2), strNewAlias(3)
       Case "Alias"
          AddAlias strNewAlias(1), strNewAlias(2), strNewAlias(3)
    End Select
    
    i = i + 1
Loop

Sub AddAlias(strAlias,strEmailAddress,strDomain)
   Dim obDomain 
   Dim obAliases 
   Dim obNewAlias

   Set obDomain = obBaseApp.Domains.ItemByName(strDomain) 
   Set obAliases = obDomain.Aliases
   Set obNewAlias = obAliases.Add() 
   
   obNewAlias.Name = strAlias & "@" & strDomain 'username
   obNewAlias.Value = strEmailAddress 'password
   obNewAlias.Active = 1 'activates user
   obNewAlias.Save() 'saves account
   
   Set obNewAlias = Nothing
   Set obAliases = Nothing
   Set obDomain = Nothing   
   
End Sub

Sub AddUser(strUsername, strPassword, strDomain)
   Dim obDomain 
   Dim obAccounts 
   Dim obNewAccount

   Set obDomain = obBaseApp.Domains.ItemByName(strDomain) 
   Set obAccounts = obDomain.Accounts
   Set obNewAccount = obAccounts.Add() 
   
   obNewAccount.Address = strUsername & "@" & strDomain 'username
   obNewAccount.Password = strPassword 'password
   obNewAccount.Active = 1 'activates user
   obNewAccount.Maxsize = 0 'sets mailbox size, 0=unlimited
   obNewAccount.Save() 'saves account
   
   Set obNewAccount = Nothing
   Set obDomain = Nothing   
   Set obAccounts = Nothing
   
End Sub