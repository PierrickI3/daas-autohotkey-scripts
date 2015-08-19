' Global variables
Public obApp
Public obAdminAccount
Public obSettings

Public Const user = "Administrator"
Public Const password = ""

' Authenticate
Set obApp = CreateObject("hMailServer.Application") 
Set obAdminAccount = obApp.Authenticate(user, password)

Set obSettings = obApp.Settings

obSettings.SecurityRanges.SetDefault()
obSettings.AutoBanOnLogonFailure = false

' Enable logging
obSettings.Logging.Enabled = true
obSettings.Logging.AWStatsEnabled = true
obSettings.Logging.LogApplication = true
obSettings.Logging.LogDebug = true
obSettings.Logging.LogIMAP = true
obSettings.Logging.LogPOP3 = true
obSettings.Logging.LogSMTP = true
obSettings.Logging.LogTCPIP = true
obSettings.Logging.KeepFilesOpen = false

' Disable Attachment blocking
obSettings.AntiVirus.EnableAttachmentBlocking = false

' Disable Spam Protection
obSettings.AntiSpam.SpamMarkThreshold = 10000
obSettings.AntiSpam.SpamDeleteThreshold = 10000
obSettings.AntiSpam.CheckHostInHelo = false
obSettings.AntiSpam.GreyListingEnabled = false
obSettings.AntiSpam.BypassGreylistingOnMailFromMX = false
obSettings.AntiSpam.SpamAssassinEnabled = false
obSettings.AntiSpam.TarpitDelay = 0
obSettings.AntiSpam.TarpitCount = 0
obSettings.AntiSpam.UseMXChecks = false
obSettings.AntiSpam.UseSPF = false
obSettings.AntiSpam.MaximumMessageSize = 1024
obSettings.AntiSpam.DKIMVerificationEnabled = false
obSettings.AntiSpam.WhiteListAddresses.Clear()
obSettings.AntiSpam.ClearGreyListingTriplets()