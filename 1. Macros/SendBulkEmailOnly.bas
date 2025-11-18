Attribute VB_Name = "Module2"
Sub SendBulkEmailOnly()
    Dim OutApp As Object, OutMail As Object, ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long, startRow As Long, endRow As Long
    Dim bccList As String, batchSize As Long, totalBatches As Long
    Dim templateFile As Variant
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    batchSize = Application.InputBox("Enter batch size (e.g., 100):", "Batch Size", 100, Type:=1)
    If batchSize <= 0 Then Exit Sub
    
    templateFile = Application.GetOpenFilename("Outlook Msg (*.msg), *.msg", , "Choose Email Template")
    If templateFile = "False" Then Exit Sub
    
    Set OutApp = CreateObject("Outlook.Application")
    
    totalBatches = Application.WorksheetFunction.RoundUp((lastRow - 1) / batchSize, 0)
    
    For i = 1 To totalBatches
        startRow = ((i - 1) * batchSize) + 2
        endRow = Application.Min(i * batchSize + 1, lastRow)
        
        bccList = ""
        For j = startRow To endRow
            If ws.Cells(j, 1).Value <> "" Then
                bccList = bccList & ws.Cells(j, 1).Value & ";"
            End If
        Next j
        
        Set OutMail = OutApp.CreateItemFromTemplate(templateFile)
        With OutMail
            .BCC = bccList
            .Display   'show draft on screen
        End With
        
        For j = startRow To endRow
            ws.Cells(j, 2).Value = "Drafted " & Format(Now, "dd-mmm-yyyy hh:mm")
        Next j
    Next i
    
    MsgBox "All batches drafted and status updated.", vbInformation
End Sub

