Attribute VB_Name = "Module3"
Sub CheckUndeliveredForSentBatch(sheetName As String, emailCol As Long, statusCol As Long)
    Dim OutApp As Object, ns As Object, inbox As Object, itm As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim mailBody As String, mailSubj As String
    Dim emailAddr As String
    
    Set ws = ThisWorkbook.Sheets(Sheet1)
    lastRow = ws.Cells(ws.Rows.Count, emailCol).End(xlUp).Row
    
    Set OutApp = CreateObject("Outlook.Application")
    Set ns = OutApp.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(6) ' Inbox
    
    For Each itm In inbox.items
        mailSubj = LCase(itm.Subject)
        
        ' Look for bounce mails
        If InStr(mailSubj, "undeliverable") > 0 Or _
           InStr(mailSubj, "delivery has failed") > 0 Or _
           InStr(mailSubj, "failure notice") > 0 Or _
           InStr(mailSubj, "returned mail") > 0 Then
           
           If itm.Class = 43 Then ' MailItem
               mailBody = itm.Body
               emailAddr = ExtractEmailFromText(mailBody)
               
               If emailAddr <> "" Then
                   For i = 2 To lastRow
                       If LCase(ws.Cells(i, emailCol).Value) = LCase(emailAddr) Then
                           ws.Cells(i, statusCol).Value = "UNDELIVERED"
                       End If
                   Next i
               End If
           End If
        End If
    Next itm
End Sub

Function ExtractEmailFromText(txt As String) As String
    Dim regEx As Object, matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[\w\.-]+@[\w\.-]+\.\w+"
    regEx.IgnoreCase = True
    regEx.Global = True
    
    If regEx.test(txt) Then
        Set matches = regEx.Execute(txt)
        ExtractEmailFromText = matches(0).Value
    Else
        ExtractEmailFromText = ""
    End If
End Function


