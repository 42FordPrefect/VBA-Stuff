Attribute VB_Name = "Modul1"
Sub ExtractLagerbestand()
    Dim fMails As Folder, mail As MailItem, txtContent As String, arrContent As Variant, objExcel As Object, wb As Object, sheet As Object, rngStart As Object, rngCurrent As Object, fErledigt As Object
    Dim PO513LLIO, PO513RLIO, PO512LLIOD, PO512RLIO, AU513LLIO, AU513RLIO, PO513LLNIO, PO513RLNIO, PO512LLNIO, PO512RLNIO, AU513LLNIO, AU513RLNIO As String
        
    'Pfad zur Excel-Datei
    Const EXCELFILE = "C:\your\file.xlsx"
    
    'Ordner in Outlook referenzieren
    Set fMails = Application.Session.Stores("your@mail.com").GetRootFolder.Folders("Ordner")
    'Unterordner referenzieren in den die Mails verschoben werden wenn sie bearbeitet wurden
    Set fErledigt = fMails.Folders("erledigt")
    
    
    If fMails.Items.Count > 0 Then
        'Excel Objekt erzeugen
        Set objExcel = CreateObject("Excel.Application")
        objExcel.DisplayAlerts = False
        
        'Excelfile öffnen
        Set wb = objExcel.Workbooks.Open(EXCELFILE)
        
        'Daten kommen in erstes Worksheet
        Set sheet = wb.Worksheets(1)
        
        'Startzelle in Spalte A ermitteln
        Set rngStart = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Offset(1, 0)
        Set rngCurrent = rngStart
        
        While fMails.Items.Count > 0
            'aktuelle Mail
            Set mail = fMails.Items(1)
            
            'Body extrahieren
            Dim oHTML As MSHTML.HTMLDocument: Set oHTML = New MSHTML.HTMLDocument
            Dim oElColl As MSHTML.IHTMLElementCollection
            With oHTML
                .Body.innerHTML = mail.HTMLBody
                Set oElColl = .getElementsByTagName("table")
                
                PO513LLIO = oElColl(0).Rows(1).Cells(1).innerText
                PO513RLIO = oElColl(0).Rows(2).Cells(1).innerText
                PO512LLIO = oElColl(0).Rows(3).Cells(1).innerText
                PO512RLIO = oElColl(0).Rows(4).Cells(1).innerText
                AU513LLIO = oElColl(0).Rows(5).Cells(1).innerText
                AU513RLIO = oElColl(0).Rows(6).Cells(1).innerText
                
                PO513LLNIO = oElColl(0).Rows(1).Cells(2).innerText
                PO513RLNIO = oElColl(0).Rows(2).Cells(2).innerText
                PO512LLNIO = oElColl(0).Rows(3).Cells(2).innerText
                PO512RLNIO = oElColl(0).Rows(4).Cells(2).innerText
                AU513LLNIO = oElColl(0).Rows(5).Cells(2).innerText
                AU513RLNIO = oElColl(0).Rows(6).Cells(2).innerText

                txt1 = oElColl(0).Rows(2).Cells(0).innerText
            End With
            txtTime = mail.ReceivedTime
            
            'Setze Werte im Sheet
            With rngCurrent
                .Value = txtStatus
                .Offset(0, 0).Value = txtTime
                
                .Offset(0, 1).Value = PO513LLIO
                .Offset(0, 2).Value = PO513RLIO
                .Offset(0, 3).Value = PO512LLIO
                .Offset(0, 4).Value = PO512RLIO
                .Offset(0, 5).Value = AU513LLIO
                .Offset(0, 6).Value = AU513RLIO
                
                .Offset(0, 7).Value = PO513LLNIO
                .Offset(0, 8).Value = PO513RLNIO
                .Offset(0, 9).Value = PO512LLNIO
                .Offset(0, 10).Value = PO512RLNIO
                .Offset(0, 11).Value = AU513LLNIO
                .Offset(0, 12).Value = AU513RLNIO
                
            End With
            'Excel Zeile eins nach unten verschieben
            Set rngCurrent = rngCurrent.Offset(1, 0)
            
            ' Mail in den 'Erledigt' Ordner verschieben
           mail.Move fErledigt
        Wend
        
        With sheet.AutoFilter.Sort
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        'Workbook speichern
        wb.Save
        'Excel anzeigen
        objExcel.Visible = True
        objExcel.DisplayAlerts = True
    Else
        MsgBox "Keine Mails zum Bearbeiten im Ordner", vbExclamation
    End If
    
    Set objExcel = Nothing
    Set wb = Nothing
    Set sheet = Nothing
    Set mail = Nothing
End Sub
