
Public Sub OpenEmailForm()
'******************************************************************
' Opens Form "Orders over $1000"
'******************************************************************

    Dim myOlApp As Application
    Dim myNameSpace As NameSpace
    Dim myFolder As MAPIFolder
    Dim myItems As Items
    Dim myItem As Object
    
    Set myOlApp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set myFolder = _
      myNameSpace.GetDefaultFolder(olFolderOutbox)
' Folder "C:\Users\user.one\AppData\Local\Microsoft\FORMS"
    
    Set myItems = myFolder.Items
    Set myItem = myItems.Add("IPM.Note.Email Form one")
    myItem.Display
    
    Set myOlApp = Nothing
    Set myNameSpace = Nothing
    Set myFolder = Nothing
    Set myItems = Nothing
    Set myItem = Nothing


End Sub
