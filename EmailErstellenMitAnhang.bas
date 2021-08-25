Attribute VB_Name = "EmailErstellenMitAnhang"
Sub Tagesbericht()
Dim mail As MailItem
Dim pfad

Dim datum As Date
Dim Anhang As Attachments


If Weekday(Date) = 2 Then
datum = Date - 3
Else
datum = Date - 1
End If


  Set mail = Application.CreateItemFromTemplate("C:\YOUR\TEMPLATE.oft")
mail.Subject = " Tagesreport " & datum
    pfad = "C:\pfad\anhang" & Format(datum, "YYMMDD") & ".pdf"

mail.Attachments.Add pfad

mail.Display


End Sub
