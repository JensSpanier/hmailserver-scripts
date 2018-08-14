Const DAYS_TO_KEEP_MESSAGES = "365"
Const MESSAGES_FOLDER = "Trash"

Const HMSADMINUSER = "Administrator"
Const HMSADMINPWD = ""

On Error Resume Next

Dim oApp
Set oApp = CreateObject("hMailServer.Application")
Call oApp.Authenticate(HMSADMINUSER, HMSADMINPWD)

For x = 0 to oApp.Domains.Count - 1
  Dim oDomain
  Set oDomain = oApp.Domains.Item(x)  
   
  For y = 0 to oDomain.Accounts.Count - 1
    Dim oAccount
    Set oAccount = oDomain.Accounts.Item(y)
   
    Dim oMessages
    Set oMessages  = oAccount.IMAPFolders.ItemByName(MESSAGES_FOLDER).Messages

    Dim DeleteMessages
    Set DeleteMessages = CreateObject("System.Collections.ArrayList")
    For z = 0 to oMessages.Count - 1
      If (DateAdd("d", DAYS_TO_KEEP_MESSAGES, CDate(oMessages.Item(z).InternalDate)) < Now ) Then
        DeleteMessages.Add oMessages.Item(z).ID
      End If
    Next
    
    For Each element In DeleteMessages
      oMessages.DeleteByDBID(element)
    Next
    
  Next
Next