Const HMSADMINUSER = "Administrator"
Const HMSADMINPWD = ""

On Error Resume Next

Dim oApp
Set oApp = CreateObject("hMailServer.Application")
Call oApp.Authenticate(HMSADMINUSER, HMSADMINPWD)

For w = 0 to oApp.Domains.Count - 1
  Dim oDomain
  Set oDomain = oApp.Domains.Item(w)  
   
  For x = 0 to oDomain.Accounts.Count - 1
    Dim oAccount
    Set oAccount = oDomain.Accounts.Item(x)
    
    CheckFolders(oAccount.IMAPFolders)
    
  Next
Next


Function CheckFolders(oFolders)
  For y = 0 to oFolders.Count - 1
    Dim oFolder
    Set oFolder = oFolders.Item(y)
    
    Dim oMessages
    Set oMessages  = oFolder.Messages
    
    Dim DeleteMessages
    Set DeleteMessages = CreateObject("System.Collections.ArrayList")
    
    For z = 0 to oMessages.Count - 1
      If (oMessages.Item(z).Flag(2)) Then
        DeleteMessages.Add oMessages.Item(z).ID
      End If
    Next
    
    For Each element In DeleteMessages
      oMessages.DeleteByDBID(element)
    Next
		   
    CheckFolders(oFolder.SubFolders)
   		
  Next
End Function