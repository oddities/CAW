Option Explicit
Dim OutlookApp, OutlookMail
Dim myNameSpace, myFolder, myNewFolder, intCount, i
Set OutlookApp = CreateObject("Outlook.Application")
Set myNameSpace = OutlookApp.GetNamespace("MAPI")
Set myFolder = myNameSpace.GetDefaultFolder(6)
'Set myNewFolder = myFolder.Folders.Add("Test")  'Create folder test
'Application.DisplayAlerts = False
For Each OutlookMail In myFolder.Items
	If OutlookMail.Subject = "Test" Then
	 'MsgBox(myFolder.Attachments)
	 'MsgBox(OutlookMail.Subject)
	 intCount = OutlookMail.Attachments.Count
     If intCount > 0 Then
        For i = 1 To intCount
            'OutlookMail.Attachments.Item(i).SaveAsFile "C:\" &  _
             MsgBox OutlookMail.Attachments.Item(i).FileName
        Next 
     End If
	End If
Next
