Attribute VB_Name = "Main"
Sub Main()

    Call CheckExpiry
    ' Set the Excel application and worksheet
    Dim excelApp As Object
    Set excelApp = ThisWorkbook.Application
    
    Dim Sh As Object
    Set Sh = ThisWorkbook.Sheets("Main")
    
    Dim Spec_Type As String
    
    Spec_Type = Sh.Cells(4, 8).Value
    
    Select Case Spec_Type
    
    Case "HDR"
        Call ImportWordDataToExcel_1
    Case "Microsoft"
        Call ImportWordDataToExcel_2
    End Select

End Sub

'Private Sub Workbook_Open()
    'Call CheckExpiry
'End Sub

'Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    'Call CheckExpiry
'End Sub

Sub CheckExpiry()
    ' Set expiry date
    Dim expiryDate As Date
    expiryDate = DateSerial(2025, 11, 1)
    
    ' Check if current date is past expiry
    If VBA.Date > expiryDate Then
        Const correctPassword As String = "Jenny0882"
        Dim userPassword As String
        userPassword = InputBox("This version has expired. Please delete this version and use the most recent version that was shared.", "Version Expired")
        
        If userPassword <> correctPassword Then
            MsgBox "Incorrect password. Please contact the administrator for the latest version.", vbCritical
            Application.DisplayAlerts = False
            ThisWorkbook.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    End If
End Sub
