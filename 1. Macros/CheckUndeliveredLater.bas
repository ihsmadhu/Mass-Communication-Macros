Attribute VB_Name = "Module3"
Option Explicit

' ==========================================================
' Procedure: CheckUndeliveredForSentBatch
'
' Description:
'   Scans the Outlook Inbox for common bounce / NDR subjects,
'   extracts the failed recipient email address from the
'   message body, and marks matching rows in an Excel sheet
'   as "UNDELIVERED".
'
' Parameters:
'   sheetName  - Name of the worksheet that contains the email list
'   emailCol   - Column number for the email addresses (e.g., 1 for column A)
'   statusCol  - Column number for the status (e.g., 2 for column B)
'
' Requirements:
'   - Outlook must be installed and configured.
'   - The target sheet must exist in ThisWorkbook.
'
' Notes:
'   - This version uses late binding for Outlook (no reference needed).
' ==========================================================
Public Sub CheckUndeliveredForSentBatch(ByVal sheetName As String, _
                                        ByVal emailCol As Long, _
                                        ByVal statusCol As Long)
    Dim OutApp As Object
    Dim ns As Object
    Dim inbox As Object
    Dim itm As Object
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim mailBody As String
    Dim mailSubj As String
    Dim emailAddr As String
    
    On Error GoTo ErrHandler
    
    ' === Excel sheet setup ===
    Set ws = ThisWorkbook.Sheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, emailCol).End(xlUp).Row
    
    ' === Outlook setup ===
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    
    If OutApp Is Nothing Then
        MsgBox "Unable to connect to Outlook.", vbCritical
        GoTo Cleanup
    End If
    
    Set ns = OutApp.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(6) ' 6 = Inbox
    
    ' === Scan Inbox items for bounce mails ===
    For Each itm In inbox.Items
        ' Guard against non-mail items
        If Not itm Is Nothing Then
            If itm.Class = 43 Then ' 43 = MailItem
                mailSubj = LCase$(NzString(itm.Subject))
                
                ' Look for bounce / NDR mails by subject
                If InStr(mailSubj, "undeliverable") > 0 Or _
                   InStr(mailSubj, "delivery has failed") > 0 Or _
                   InStr(mailSubj, "failure notice") > 0 Or _
                   InStr(mailSubj, "returned mail") > 0 Then
                   
                    mailBody = NzString(itm.Body)
                    emailAddr = ExtractEmailFromText(mailBody)
                    
                    If emailAddr <> "" Then
                        ' Match against email list in the sheet
                        For i = 2 To lastRow   ' Assuming row 1 is header
                            If LCase$(Trim$(ws.Cells(i, emailCol).Value)) = LCase$(Trim$(emailAddr)) Then
                                ws.Cells(i, statusCol).Value = "UNDELIVERED"
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    Next itm
    
Cleanup:
    Set itm = Nothing
    Set inbox = Nothing
    Set ns = Nothing
    Set OutApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error in CheckUndeliveredForSentBatch: " & Err.Number & _
           " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================================
' Helper: Extract the first email address found in a text
' ==========================================================
Public Function ExtractEmailFromText(ByVal txt As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Pattern = "[\w\.-]+@[\w\.-]+\.\w+"
        .IgnoreCase = True
        .Global = True
    End With
    
    If regEx.test(txt) Then
        Set matches = regEx.Execute(txt)
        ExtractEmailFromText = matches(0).Value
    Else
        ExtractEmailFromText = ""
    End If
End Function

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
