Attribute VB_Name = "Module2"
Option Explicit

' ==========================================================
' Procedure: SendBulkEmailOnly
'
' Description:
'   - Reads email addresses from an Excel sheet.
'   - Sends them in batches as BCC recipients using an
'     Outlook .msg template file.
'   - For each batch, creates a draft email and updates
'     the status column in Excel.
'
' Behavior:
'   - Prompts the user for:
'       * Batch size (e.g., 100 recipients per email).
'       * An Outlook .msg file to use as the email template.
'   - For each batch:
'       * Builds a semicolon-separated BCC list.
'       * Creates a mail item from the template.
'       * Sets the BCC field.
'       * Displays the draft (user can review/send manually).
'       * Marks all rows in that batch as "Drafted <timestamp>".
'
' Requirements:
'   - Excel workbook with a sheet containing:
'       * Column A: Email addresses
'       * Column B: Status (will be updated by the macro)
'   - Outlook installed and configured.
'
' Notes:
'   - Uses late binding for Outlook (no reference needed).
'   - Adjust SHEET_NAME / COL_EMAIL / COL_STATUS as needed.
'
' Author: <Your Name>
' GitHub: https://github.com/<your-handle>/<your-repo>
' ==========================================================

Private Const SHEET_NAME As String = "Sheet1"
Private Const COL_EMAIL As Long = 1   ' Column A
Private Const COL_STATUS As Long = 2  ' Column B

Public Sub SendBulkEmailOnly()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim startRow As Long
    Dim endRow As Long
    
    Dim bccList As String
    Dim batchSize As Long
    Dim totalBatches As Long
    Dim templateFile As Variant
    
    On Error GoTo ErrHandler
    
    ' === Excel sheet setup ===
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_EMAIL).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No email addresses found in " & SHEET_NAME & ".", vbExclamation
        GoTo Cleanup
    End If
    
    ' === Ask for batch size ===
    batchSize = Application.InputBox( _
                    Prompt:="Enter batch size (e.g., 100):", _
                    Title:="Batch Size", _
                    Default:=100, _
                    Type:=1)
    ' If user cancels or enters invalid value
    If batchSize <= 0 Then GoTo Cleanup
    
    ' === Ask for Outlook .msg template ===
    templateFile = Application.GetOpenFilename( _
                        FileFilter:="Outlook Msg (*.msg), *.msg", _
                        Title:="Choose Email Template")
    ' User clicked cancel
    If templateFile = False Then GoTo Cleanup
    
    ' === Outlook application ===
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    
    If OutApp Is Nothing Then
        MsgBox "Unable to start Outlook.", vbCritical
        GoTo Cleanup
    End If
    
    ' === Calculate number of batches ===
    totalBatches = Application.WorksheetFunction.RoundUp((lastRow - 1) / batchSize, 0)
    
    ' === Loop through batches ===
    For i = 1 To totalBatches
        startRow = ((i - 1) * batchSize) + 2      ' row 2 is first email (row 1 = header)
        endRow = Application.Min(i * batchSize + 1, lastRow)
        
        ' Build BCC list for the current batch
        bccList = ""
        For j = startRow To endRow
            If NzString(ws.Cells(j, COL_EMAIL).Value) <> "" Then
                bccList = bccList & ws.Cells(j, COL_EMAIL).Value & ";"
            End If
        Next j
        
        ' If no emails in this batch, skip
        If bccList <> "" Then
            ' Create draft from template
            Set OutMail = OutApp.CreateItemFromTemplate(templateFile)
            
            With OutMail
                .BCC = bccList
                .Display   ' show the draft on screen; user sends manually
            End With
            
            ' Mark rows as drafted
            For j = startRow To endRow
                If NzString(ws.Cells(j, COL_EMAIL).Value) <> "" Then
                    ws.Cells(j, COL_STATUS).Value = _
                        "Drafted " & Format(Now, "dd-mmm-yyyy hh:mm")
                End If
            Next j
        End If
    Next i
    
    MsgBox "All batches drafted and status updated.", vbInformation

Cleanup:
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error in SendBulkEmailOnly: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================================
' Helper: Null/Empty-safe string
' ==========================================================
Private Function NzString(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NzString = ""
    Else
        NzString = CStr(value)
    End If
End Function
