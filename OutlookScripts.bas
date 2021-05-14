'This script will allow the user to choose which folder to save the sent message, similar to the Lotus Notes features.
'If the User selects "cancel", then the email will be saved to the Sent folder automatically.

'To implement this script:
' - Copy this whole script
' - In Outlook, press ALT+F11 to open the VB Editor
' - Use the left navigation pane to navigate to "ThisOutlookSession" and double click it.
' - Paste the contents into the large space.
' - Save
' - Close the VB Editor

Private Sub Application_ItemSend(ByVal Item As Object, _
    Cancel As Boolean)
  Dim objNS As NameSpace
  Dim objFolder As MAPIFolder
  Set objNS = Application.GetNamespace("MAPI")
  Set objFolder = objNS.PickFolder
  If TypeName(objFolder) <> "Nothing" And _
     IsInDefaultStore(objFolder) Then
      Set Item.SaveSentMessageFolder = objFolder
  End If
  Set objFolder = Nothing
  Set objNS = Nothing
End Sub

Public Function IsInDefaultStore(objOL As Object) As Boolean
  Dim objApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Dim objInbox As Outlook.MAPIFolder
  On Error Resume Next
  Set objApp = CreateObject("Outlook.Application")
  Set objNS = objApp.GetNamespace("MAPI")
  Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
  Select Case objOL.Class
    Case olFolder
      If objOL.StoreID = objInbox.StoreID Then
        IsInDefaultStore = True
      End If
    Case olAppointment, olContact, olDistributionList, _
         olJournal, olMail, olNote, olPost, olTask
      If objOL.Parent.StoreID = objInbox.StoreID Then
        IsInDefaultStore = True
      End If
    Case Else
      MsgBox "This function isn't designed to work " & _
             "with " & TypeName(objOL) & _
             " items and will return False.", _
             , "IsInDefaultStore"
  End Select
  Set objApp = Nothing
  Set objNS = Nothing
  Set objInbox = Nothing
End Function
