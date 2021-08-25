Attribute VB_Name = "MailZuAufgabe"
Sub EmailZuAufagabe()
Dim mail As MailItem
Dim Aufgabe As TaskItem
Set mail = Application.ActiveExplorer.Selection.Item(1)
If TypeName(mail) = "MailItem" Then
    Set Aufgabe = Application.CreateItem(olTaskItem)
    Aufgabe.Body = mail.Body
    Aufgabe.Subject = mail.Subject
    Aufgabe.Attachments.Add mail
    Aufgabe.StartDate = Now()
    Aufgabe.Save
    
Else
    MsgBox ("Bitte Mail Auswählen")
End If
    

End Sub
