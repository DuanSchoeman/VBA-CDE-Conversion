Attribute VB_Name = "Extract_Data"

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr

Sub ImportWordDataToExcel_1()
    
    ' Set the Excel application and worksheet
    Dim excelApp As Object
    Set excelApp = ThisWorkbook.Application ' Assumes the code is in the Excel workbook
      
    
    Dim startTime As Double
    Dim endTime As Double
    Dim duration As Double
    
    ' Record the start time
    startTime = Timer
    
    Dim savedEnableEvents As Boolean
    savedEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim filePath As String
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    ' Configure the file dialog
    With fileDialog
        .InitialFileName = GetWorkbookPath()
        .AllowMultiSelect = False
        .Title = "Select a Word file"
        .Filters.Clear
        .Filters.Add "Word and Rich Text Files", "*.docx; *.doc; *.rtf"
    End With
    
    ' Show the file dialog
    If fileDialog.Show = -1 Then
        filePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "No file selected"
        Exit Sub
    End If
        
    Dim excelSheet As Object
    Dim FileName As String
    Dim MainExcelSheet As Object
    ' Extract document name from file path
    FileName = ExtractDocumentName(filePath)

    ' Set the source sheet to copy from
    Dim sourceSheet As Object
    Set sourceSheet = ThisWorkbook.Sheets("Reference_Sheet_HDR")
    Set MainExcelSheet = ThisWorkbook.Sheets("Main")
    ' Copy the source sheet
    sourceSheet.Copy After:=Sheets(Sheets.Count)
    ' Set the new sheet
    Set excelSheet = ActiveSheet
    ' Rename the new sheet
    If Len(FileName) > 31 Then
        excelSheet.Name = Left(FileName, 31)
    Else
        excelSheet.Name = FileName
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    Dim lastRow As Long
    lastRow = excelSheet.Cells(excelSheet.Rows.Count, "B").End(xlUp).Row

    If lastRow >= 9 Then
        
        ' Unhide rows within the range
        excelSheet.Rows("9:" & lastRow).EntireRow.Hidden = False

        ' Delete the range, but keep rows 10 and 11
        If lastRow > 9 Then
            excelSheet.Rows("10:" & lastRow).Delete
        End If
    End If

    excelSheet.Cells(9, 2).Value = FileName
    excelSheet.Cells(1, 2).Value = UCase(FileName)
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim objDoc As Object
    Dim docOpened As Boolean
    
    ' Check if there is an existing instance of Word open
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo 0
    
    If Not wordApp Is Nothing Then
        ' Check if the desired document is already open
        For Each objDoc In wordApp.Documents
            If objDoc.FullName = filePath Then
                ' Document is already open, use the existing instance
                Set wordDoc = objDoc
                docOpened = True
                Exit For
            End If
        Next objDoc
    Else
        Set wordApp = CreateObject("Word.Application")
    End If
    
    ' If the document is not open, open it
    If Not docOpened Then
        Set wordDoc = wordApp.Documents.Open(filePath)
    End If
    
    RemoveHeadersAndFooters wordDoc

    ' Set the starting row index
    Dim startingRowIndex As Integer
    startingRowIndex = 1
    
    Dim numLines As Integer
    Dim cleanedText As String
    Dim Starting_Nums() As Variant
    Dim Ending_Nums() As Variant
    Dim First_Heading As Boolean
    Dim Count As Integer
    Dim End_Count As Integer
    First_Heading = False
    Count = 1
    End_Count = 1
    Dim Appendix As Boolean
    Dim pastedTables As Object
    Dim pastedTables_Create As Object
    Set pastedTables = CreateObject("Scripting.Dictionary")
    Set pastedTables_Create = CreateObject("Scripting.Dictionary")
    Set PastedImage = CreateObject("Scripting.Dictionary")
    Set PastedImage_Inline = CreateObject("Scripting.Dictionary")
    Dim ErrorCounter As Integer
    Dim ErrorCounter_IMG As Integer
    Dim tableHandled As Boolean
    Dim currentPageNumber As Long
    Dim maxColumnCount As Integer
    maxColumnCount = 0
    Dim currentPageNumber_Create As Integer
    Dim textWidth As Double
    Dim lineHeight As Double
    Dim Max_Char_Cnt As Integer
    Dim App_A1_isTrue As Boolean
    Dim Tab_parts() As String
    Dim Heading_Size_Min As Integer
    Dim Heading_Size_Max As Integer
    Dim spacePos As Integer
    Dim tabPos As Integer
    Dim firstDelimiterPos As Integer
    Dim headingParts() As String
    Dim paraIndex As Integer
    Dim Table_Image As Boolean
    Dim Group_Headings As Boolean
    
    paraIndex = 1

    Dim detectedSizes As Variant
    detectedSizes = AutoDetectHeadingSizes(wordDoc)
    
    Heading_Size_Min = detectedSizes(1)
    Heading_Size_Max = detectedSizes(2)
    
    Table_Image = MainExcelSheet.Cells(5, 8).Value
    Group_Headings = MainExcelSheet.Cells(3, 8).Value
    
    If Heading_Size_Min = 0 Then
        MsgBox "The textsize of the paragraph was not set, a size of 10 will be used."
        Heading_Size_Min = 10
    End If
    
    App_A1_isTrue = False
    
    lineHeight = 11 * 1.6
    
    ' Create a new sheet to paste the tables
    Dim excelSheet_Create As Object
    Set excelSheet_Create = ThisWorkbook.Sheets.Add

    
    If Table_Image = False Then
        ' Iterate through tables in the Word document
        For Each tbl In wordDoc.Tables
            On Error Resume Next
                ' Get the page number of the table
                currentPageNumber_Create = tbl.Range.Information(wdActiveEndPageNumber)
                
                ' Table Key logic
                Dim tableKey_Counter As String
                tableKey_Counter = tbl.Range.Start & ":" & tbl.Range.End & ":" & currentPageNumber_Create
                If pastedTables_Create.Exists(tableKey_Counter) Then
                    ' Table has already been pasted, skip
                    GoTo NextTable
                End If
                
                ' Copy the table from Word
                tbl.Range.Copy
                
                ' Select the destination cell in Excel
                excelSheet_Create.Cells(startingRowIndex, 1).Select
                
                ' Paste the table in Excel
                excelSheet_Create.Paste
                
                ' Clear the clipboard
                Application.CutCopyMode = False
            
                ' Count the columns in the pasted table
                Dim pastedColumnCount As Integer
                pastedColumnCount = excelSheet_Create.UsedRange.Columns.Count
                
                ' Clear the pasted columns
                excelSheet_Create.Columns(startingRowIndex).Resize(, pastedColumnCount).Delete 'Shift:=xlToLeft
            
                ' Update the maximum column count
                If pastedColumnCount > maxColumnCount Then
                    maxColumnCount = pastedColumnCount
                End If
            
                ' Increment the starting row index by the number of rows in the pasted table
                startingRowIndex = startingRowIndex + 2 * (tbl.Rows.Count)
            
                ' Mark the table as pasted
                pastedTables_Create.Add tableKey_Counter, True
            On Error GoTo 0
NextTable:
        Next tbl
    Else
        maxColumnCount = 8
    End If
    ' Delete the sheet without onscreen messages
    
    excelSheet_Create.Delete
    
    'Where the first excel entry will be
    startingRowIndex = 10

    If maxColumnCount > 8 Then
        Dim i As Integer
        For i = 1 To (maxColumnCount - 8)
            excelSheet.Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Next i
    Else
        maxColumnCount = 8
    End If
    
    Max_Char_Cnt = 26 * maxColumnCount
    
    For Each para In wordDoc.Paragraphs
    
        If para.Range.Information(3) <> currentPageNumber Then ' Assuming 3 corresponds to wdActiveEndPageNumber
            ' Update the page number
            currentPageNumber = para.Range.Information(3)
        End If
        
        Appendix = False
        
       
        If (Not para.Range.Tables.Count > 0) And ((para.Style Like "Heading*" And Len(TrimSpacesAndTabs(para.Range.text)) > 0) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And IsEntireTextBold(para.Range) And para.Range.Font.size >= Heading_Size_Min And para.Style <> "List Paragraph" And para.Range.Tables.Count = 0) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextBold(para.Range) And para.Range.Font.size >= Heading_Size_Min) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Range.Font.Bold = False And para.Style = "List Paragraph" And (UBound(Split(para.Range.ListFormat.ListString, ".")) > 1) And para.Range.Font.size >= Heading_Size_Min)) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextCapitals(para.Range) And IsNumeric(Left(para.Range.ListFormat.ListString, 1)) And para.Range.Tables.Count = 0) Then

            
            ' Extract heading number and text
            Dim headingNumber As String
            Dim Heading_Level As Integer
            Dim headingText As String
            
            If Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style Like "Heading*" Then
                headingNumber = para.Range.ListFormat.ListString
                Heading_Level = para.OutlineLevel
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Range.Font.size >= Heading_Size_Min And IsEntireTextBold(para.Range) And para.Style <> "List Paragraph" Then
                ' Find the position of the first space and the first tab
                spacePos = InStr(para.Range.text, " ")
                tabPos = InStr(para.Range.text, vbTab)
                
                ' Determine which comes first (ignoring zeroes)
                If spacePos > 0 And tabPos > 0 Then
                    firstDelimiterPos = WorksheetFunction.Min(spacePos, tabPos)
                ElseIf spacePos > 0 Then
                    firstDelimiterPos = spacePos
                ElseIf tabPos > 0 Then
                    firstDelimiterPos = tabPos
                Else
                    firstDelimiterPos = 0
                End If
                
                If firstDelimiterPos >= 1 Then
                    headingNumber = Left(para.Range.text, firstDelimiterPos - 1)
                Else
                    headingNumber = ""
                End If

                
                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")
                
                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And para.Range.Font.size >= Heading_Size_Min And IsEntireTextBold(para.Range) Then
                headingNumber = para.Range.ListFormat.ListString
                Heading_Level = para.Range.ListFormat.ListLevelNumber
            ElseIf (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Range.Font.Bold = False And para.Style = "List Paragraph" And (UBound(Split(para.Range.ListFormat.ListString, ".")) > 1) And para.Range.Font.size >= Heading_Size_Min) Then
                headingNumber = para.Range.ListFormat.ListString

                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")

                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then ' And para.Range.Font.Size >= Heading_Size_Min
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextCapitals(para.Range) And IsNumeric(Left(para.Range.ListFormat.ListString, 1)) And para.Range.Tables.Count = 0 Then
                headingNumber = para.Range.ListFormat.ListString

                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")

                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then ' And para.Range.Font.Size >= Heading_Size_Min
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            End If
            
            headingNumber = para.Range.ListFormat.ListString
            headingText = para.Range.text
            
            If tableHandled Then
                ' Move to the next paragraph
                Set para = para.Next
                ' Reset the flag
                tableHandled = False
            End If
            
            

            ' Check if it's the highest level heading
            If Len(headingNumber) > 0 Then
                If Heading_Level = 1 Then
                    
                    ReDim Preserve Starting_Nums(0 To Count - 1)
                    Starting_Nums(Count - 1) = startingRowIndex + 1
                    
                    If First_Heading = True Then
                        ReDim Preserve Ending_Nums(0 To Count - 2)
                        Ending_Nums(Count - 2) = startingRowIndex - 1
                        End_Count = End_Count + 1
                    End If
                    
                    Count = Count + 1
                    First_Heading = True
                    
                    
                    ' Merge, format, and add borders to the merged cell for top-level heading
                    With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)) '14
                        .Merge
                        .HorizontalAlignment = -4131 ' Left align
                        .VerticalAlignment = -4108 ' Center vertically
                        .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                        .Borders.LineStyle = xlContinuous ' Set outside borders
                    End With
    
                    ' Set formatting for top-level heading
                    With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font '14
                        .Name = "Calibri" ' Font name
                        .size = 16 ' Font size
                        .Bold = True ' Bold
                    End With
                    
                Else
    
                    If Heading_Level = 2 Then
                        ' Second-level heading
                        ' Your code for second-level formatting
                        ' Merge, format, and add borders to the merged cell for second-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                            .Merge
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(221, 217, 196) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
                        ' Set formatting for second-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                            .Name = "Calibri" ' Font name
                            .size = 12 ' Font size
                            .Bold = True ' Bold
                        End With
                    ElseIf Heading_Level = 3 Then
                        ' Third-level heading
                        ' Your code for third-level formatting
                        ' Merge, format, and add borders to the merged cell for third-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                            .Merge
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(221, 217, 196) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
    
                        ' Set formatting for third-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                            .Name = "Calibri" ' Font name
                            .size = 11 ' Font size
                            .Bold = True ' Bold
                        End With
                    ElseIf Heading_Level = 4 Then
                        ' Fourth-level heading
                        ' Your code for third-level formatting
                        ' Merge, format, and add borders to the merged cell for third-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                            .Merge
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(245, 240, 220) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
    
                        ' Set formatting for third-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                            .Name = "Calibri" ' Font name
                            .size = 11 ' Font size
                            .Bold = False ' Bold
                        End With
                        
                    Else
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                            .Merge
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
        
                        ' Set formatting for top-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                            .Name = "Calibri" ' Font name
                            .size = 16 ' Font size
                            .Bold = True ' Bold
                        End With
                        
                    End If
                End If
                
            Else
                
                Dim appendixParts() As String
                
                If ((InStr(1, headingText, "APPENDIX") > 0) Or (InStr(1, headingText, "Appendix") > 0)) And ((InStr(1, headingText, "Publications Referenced")) = 0) Then
                   Appendix = True
                   
                   ' Split the heading into two parts based on the tab character
                    appendixParts = Split(headingText, vbTab)
                   ' Merge, format, and add borders to the merged cell for top-level heading
                   With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                       .Merge
                       .HorizontalAlignment = -4131 ' Left align
                       .VerticalAlignment = -4108 ' Center vertically
                       .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                       .Borders.LineStyle = xlContinuous ' Set outside borders
                   End With
    
                   ' Set formatting for top-level heading
                   With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                       .Name = "Calibri" ' Font name
                       .size = 16 ' Font size
                       .Bold = True ' Bold
                   End With


                   ReDim Preserve Starting_Nums(0 To Count - 1)
                   Starting_Nums(Count - 1) = startingRowIndex + 1

                   ReDim Preserve Ending_Nums(0 To Count - 2)
                   Ending_Nums(Count - 2) = startingRowIndex - 1
                   End_Count = End_Count + 1

                   Count = Count + 1
                ElseIf (InStr(1, headingText, "Table") > 0) Or (InStr(1, headingNumber, "Table") > 0) Then
                
                    With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount))
                        .Merge
                        .HorizontalAlignment = -4131 ' Left align
                        .VerticalAlignment = -4108 ' Center vertically
                    End With
                    
                    ' Set formatting for third-level heading
                    With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount).Font
                        .Name = "Calibri" ' Font name
                        .size = 12 ' Font size
                        .Bold = True ' Bold
                    End With
                    
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Name = "Times New Roman"
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.size = 11
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)).Borders.LineStyle = xlContinuous
                    
                ElseIf ((InStr(1, headingText, "APPENDIX") > 0) And (InStr(1, headingText, "Publications Referenced")) > 0) Then
                
                    App_A1_isTrue = True
                    Appendix = True
                   
                   ' Split the heading into two parts based on the tab character
                    appendixParts = Split(headingText, vbTab)
                   ' Merge, format, and add borders to the merged cell for top-level heading
                   With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                       .Merge
                       .HorizontalAlignment = -4131 ' Left align
                       .VerticalAlignment = -4108 ' Center vertically
                       .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                       .Borders.LineStyle = xlContinuous ' Set outside borders
                   End With
    
                   ' Set formatting for top-level heading
                   With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                       .Name = "Calibri" ' Font name
                       .size = 16 ' Font size
                       .Bold = True ' Bold
                   End With


                   ReDim Preserve Starting_Nums(0 To Count - 1)
                   Starting_Nums(Count - 1) = startingRowIndex + 1

                   ReDim Preserve Ending_Nums(0 To Count - 2)
                   Ending_Nums(Count - 2) = startingRowIndex - 1
                   End_Count = End_Count + 1

                   Count = Count + 1
                Else

                    If Heading_Level = 1 Then
                        
                        ReDim Preserve Starting_Nums(0 To Count - 1)
                        Starting_Nums(Count - 1) = startingRowIndex + 1
                        
                        If First_Heading = True Then
                            ReDim Preserve Ending_Nums(0 To Count - 2)
                            Ending_Nums(Count - 2) = startingRowIndex - 1
                            End_Count = End_Count + 1
                        End If
                        
                        Count = Count + 1
                        First_Heading = True
                        
                        
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                            .Merge
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
        
                        ' Set formatting for top-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                            .Name = "Calibri" ' Font name
                            .size = 16 ' Font size
                            .Bold = True ' Bold
                        End With
                        
                    Else
        
                        If Heading_Level = 2 Then
                            ' Second-level heading
                            ' Your code for second-level formatting
                            ' Merge, format, and add borders to the merged cell for second-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                                .Merge
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(221, 217, 196) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
                            ' Set formatting for second-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                                .Name = "Calibri" ' Font name
                                .size = 12 ' Font size
                                .Bold = True ' Bold
                            End With
                        ElseIf Heading_Level = 3 Then
                            ' Third-level heading
                            ' Your code for third-level formatting
                            ' Merge, format, and add borders to the merged cell for third-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                                .Merge
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(221, 217, 196) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
        
                            ' Set formatting for third-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                                .Name = "Calibri" ' Font name
                                .size = 11 ' Font size
                                .Bold = True ' Bold
                            End With
                        ElseIf Heading_Level = 4 Then
                            ' Fourth-level heading
                            ' Your code for third-level formatting
                            ' Merge, format, and add borders to the merged cell for third-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                                .Merge
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(245, 240, 220) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
        
                            ' Set formatting for third-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                                .Name = "Calibri" ' Font name
                                .size = 11 ' Font size
                                .Bold = False ' Bold
                            End With
                            
                        Else
                            ' Merge, format, and add borders to the merged cell for top-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                                .Merge
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(196, 189, 151) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
            
                            ' Set formatting for top-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, 1 + maxColumnCount + 4).Font
                                .Name = "Calibri" ' Font name
                                .size = 16 ' Font size
                                .Bold = True ' Bold
                            End With
                            
                        End If
                    End If
                
                End If
            End If
'TrimSpacesAndTabs
            If TrimSpacesAndTabs(headingText) <> "" Then  'Trim(headingText)
                ' Set the value for the merged cell
                If Appendix = False Then
                    If headingNumber <> "" Then
                        excelSheet.Cells(startingRowIndex, 1).Value = headingNumber & Space(34 - Len(headingNumber)) & TrimSpacesAndTabs(headingText)
                    Else
                        If Len(TrimSpacesAndTabs(headingNumber)) > 0 Then
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(headingNumber) & Space(34 - Len(headingNumber)) & TrimSpacesAndTabs(headingText)
                        Else
                            excelSheet.Cells(startingRowIndex, 1).Value = Space(10) & TrimSpacesAndTabs(headingText)
                        End If
                    End If
                Else
                    If UBound(appendixParts) > 0 Then
                        excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(appendixParts(0)) & Space(21 - Len(TrimSpacesAndTabs(appendixParts(0)))) & TrimSpacesAndTabs(appendixParts(1))
                    Else
                        excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(appendixParts(0))
                    End If
                End If
            End If
            
            ' Increment the starting row index
            startingRowIndex = startingRowIndex + 1
                

        ElseIf para.Range.Tables.Count > 0 Then
            
            Set tbl = para.Range.Tables(1)
            Dim tableKey As String
            tableKey = para.Range.Tables(1).Range.Start & ":" & para.Range.Tables(1).Range.End & ":" & currentPageNumber
            
            If Not pastedTables.Exists(tableKey) Then
                If Table_Image Then
                    ' Handle table as image
                    Dim tblWidth As Double
                    Dim tblHeight As Double
                    
                    
                    ' Copy again
                    tbl.Range.CopyAsPicture
                    Application.Wait Now + TimeValue("00:00:01")
                    DoEvents
                    
                    On Error Resume Next
                    
                    excelSheet.Cells(startingRowIndex, 2).PasteSpecial Paste:=xlPasteEnhancedMetafile ', Link:=False, DisplayAsIcon:=False
                    
                    If Err.Number <> 0 Then
                        
                        Err.Clear
                        ErrorCounter = ErrorCounter + 1
                        'GoTo RegularTablePasting Disabled because column count hasn't been adjusted to accomodate tables that have more columns than what is available on the template.
                    End If
                    
                    On Error GoTo 0
                    
                    ' Get the pasted image shape
                    Dim pastedShape As Shape
                    Set pastedShape = excelSheet.Shapes(excelSheet.Shapes.Count)
                    
                    If Not pastedShape Is Nothing Then
                        tblWidth = pastedShape.Width
                        tblHeight = pastedShape.Height
                        
                        Dim cellWidthInPoints As Double
                        Dim cellHeightInPoints As Double
                        cellWidthInPoints = excelSheet.Cells(startingRowIndex, 2).Width
                        cellHeightInPoints = excelSheet.Cells(startingRowIndex, 2).Height
                        
                        Dim Cell_Horiz_Cnt As Integer
                        Dim Cell_Vert_Cnt As Integer
                        Cell_Horiz_Cnt = WorksheetFunction.RoundUp(tblWidth / cellWidthInPoints, 0)
                        Cell_Vert_Cnt = WorksheetFunction.RoundUp(tblHeight / cellHeightInPoints, 0)
                        
                        Dim ScaleFactor As Double
                        ScaleFactor = 1
                        If Cell_Horiz_Cnt > maxColumnCount Then
                            ScaleFactor = (Cell_Horiz_Cnt / maxColumnCount) - 1
                            Cell_Horiz_Cnt = maxColumnCount
                            
                            pastedShape.Height = cellHeightInPoints * Cell_Vert_Cnt * ScaleFactor
                            pastedShape.Width = cellWidthInPoints * Cell_Horiz_Cnt * ScaleFactor
                            pastedShape.Top = excelSheet.Cells(startingRowIndex, 2).Top
                            pastedShape.Left = excelSheet.Cells(startingRowIndex, 2).Left
                        End If
                        
                        pastedShape.Placement = xlMoveAndSize ' Ensure the image is moved and sized with cells
                            
                        
                        ' Add borders
                        excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, 1)).Borders.LineStyle = xlContinuous
                        excelSheet.Range(excelSheet.Cells(startingRowIndex, 2 + maxColumnCount), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, 1 + maxColumnCount + 4)).Borders.LineStyle = xlContinuous
                        
                        ' Update the starting row index
                        startingRowIndex = startingRowIndex + Cell_Vert_Cnt + 1
                    Else
                        'GoTo RegularTablePasting Disabled because column count hasn't been adjusted to accomodate tables that have more columns than what is available on the template.
                    End If
                Else
RegularTablePasting:
                    ' Original table handling code
                    ' Set the Excel range where you want to paste the table
                    Dim excelTableRange As Object
                    Set excelTableRange = excelSheet.Range("B" & startingRowIndex)
            
                    ' Copy the table from Word
                    tbl.Range.Copy
                    Application.Wait Now + TimeValue("00:00:01")
                    ' Select the destination cell in Excel
                    excelTableRange.Select
                    
                    On Error Resume Next
                    ' Paste the table in Excel
                    excelSheet.Paste
            
                    ' Clear the clipboard
                    Application.CutCopyMode = False
            
                    ' Increment the starting row index
                    Dim Table_lastRow(1 To 8) As Long
                    'Dim i As Integer
                    For i = 1 To 8
                        Table_lastRow(i) = excelSheet.Cells(excelSheet.Rows.Count, i + 1).End(xlUp).Row
                    Next i
                    
                    Dim maxRows As Long
                    maxRows = WorksheetFunction.Max(Table_lastRow(1), Table_lastRow(2), Table_lastRow(3), Table_lastRow(4), _
                                                    Table_lastRow(5), Table_lastRow(6), Table_lastRow(7), Table_lastRow(8))
                                        
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(maxRows, 1 + maxColumnCount)).HorizontalAlignment = xlCenter
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(maxRows, 1 + maxColumnCount)).VerticalAlignment = xlCenter
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 1), excelSheet.Cells(maxRows, 1 + maxColumnCount + 4)).Borders.LineStyle = xlContinuous
                    
                    ' Merge cells in column A corresponding to the table rows
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(maxRows, 1)).Merge
                    
                    ' Add borders to the merged cell in column A
                    With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(maxRows, 1))
                        .Borders.LineStyle = xlContinuous
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    startingRowIndex = maxRows + 1
                    
                    If Err.Number <> 0 Then
                        ' An error occurred, so increment the error counter
                        ErrorCounter = ErrorCounter + 1
                        ' Reset the error
                        Err.Clear
                        
                        startingRowIndex = startingRowIndex + 2
                        
                        ' Write "Failed Table Space"
                        excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                        excelSheet.Cells(startingRowIndex, 2).Value = "Placeholder - Failed To Copy Table"
                        excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)).Interior.Color = RGB(255, 255, 0)
                        
                        ' Increment the starting row index
                        startingRowIndex = startingRowIndex + 2
                    End If
                End If
                
                ' Mark the table as pasted
                pastedTables.Add tableKey, True
            End If
            
            On Error GoTo 0
            
        ElseIf para.Range.ListFormat.ListType <> wdListNoNumbering Then
       
            
            Dim listLevel As Integer
            Dim Bullet_Char As String
            
            listLevel = para.Range.ListFormat.ListLevelNumber
            
            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)
            
            If listLevel <= 1 Then
                Bullet_Char = ChrW(&H2022)
            ElseIf (listLevel > 1) And (listLevel <= 2) Then
                Bullet_Char = ChrW(&H25CB)
            ElseIf (listLevel > 2) And (listLevel <= 3) Then
                Bullet_Char = ChrW(&H25C7)
            Else
                Bullet_Char = para.Range.ListFormat.ListString
            End If
                        
            Dim firstChar As String
            firstChar = Left(para.Range.ListFormat.ListString, 1)
            
            
            If para.Range.ListFormat.ListType = wdListSimpleNumbering Or para.Range.ListFormat.ListType = wdListOutlineNumbering Then
                
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, maxColumnCount - 1)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = para.Range.ListFormat.ListString & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, maxColumnCount).Value = TrimSpacesAndTabs(Tab_parts(1))
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 3)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 4), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = para.Range.ListFormat.ListString & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, 4).Value = TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                    excelSheet.Cells(startingRowIndex, 2).Value = para.Range.ListFormat.ListString & " " & cleanedText
                End If
                excelSheet.Cells(startingRowIndex, 2).IndentLevel = listLevel
            
            ElseIf IsNumeric(firstChar) And Mid(para.Range.ListFormat.ListString, 3, 1) <> "." Then
                'Paste just the content in Excel
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, maxColumnCount - 1)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, maxColumnCount).Value = TrimSpacesAndTabs(Tab_parts(1))
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 3)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 4), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, 4).Value = TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                    excelSheet.Cells(startingRowIndex, 2).Value = cleanedText
                End If
                    'excelSheet.Cells(startingRowIndex, 2).Value = cleanedText
            

            ElseIf para.Range.ListFormat.ListType = wdListBullet Then
            
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, maxColumnCount - 1)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, maxColumnCount).Value = TrimSpacesAndTabs(Tab_parts(1))
                    'excelSheet.Cells(startingRowIndex, 2).Value = para.Range.ListFormat.ListString & " " & cleanedText
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 4)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 5), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, 5).Value = TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & cleanedText
                End If

                excelSheet.Cells(startingRowIndex, 2).IndentLevel = listLevel
            Else
                
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, maxColumnCount - 1)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, maxColumnCount).Value = TrimSpacesAndTabs(Tab_parts(1))
                    'excelSheet.Cells(startingRowIndex, 2).Value = para.Range.ListFormat.ListString & " " & cleanedText
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Merge columns from B to I
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 4)).Merge
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 5), excelSheet.Cells(startingRowIndex, maxColumnCount + 1)).Merge
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0))
                    excelSheet.Cells(startingRowIndex, 5).Value = TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                    excelSheet.Cells(startingRowIndex, 2).Value = Bullet_Char & " " & cleanedText
                End If

                excelSheet.Cells(startingRowIndex, 2).IndentLevel = listLevel
            End If
            
            excelSheet.Rows(startingRowIndex).WrapText = True
            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)).Borders.LineStyle = xlContinuous
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Name = "Times New Roman"
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.size = 11
            
            If para.Range.Font.Underline <> wdUnderlineNone Then
                excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Underline = True
            End If
            
            ' Merge, format, and add borders to the merged cell for top-level heading
            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)) '14
                .HorizontalAlignment = -4131 ' Left align
                .VerticalAlignment = -4108 ' Center vertically
            End With
            
            
            startingRowIndex = startingRowIndex + 1
            
        ElseIf para.Style = "Table Caption" Then

            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)

            excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
            excelSheet.Cells(startingRowIndex, 2).Value = cleanedText
            excelSheet.Rows(startingRowIndex).WrapText = True
            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)).Borders.LineStyle = xlContinuous
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Name = "Times New Roman"
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.size = 11
            
            ' Merge, format, and add borders to the merged cell for top-level heading
            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4))
                .HorizontalAlignment = -4131 ' Left align
                .VerticalAlignment = -4108 ' Center vertically
            End With
            
            ' Calculate the number of lines in the paragraph
            startingRowIndex = startingRowIndex + 1
            
        ElseIf para.Range.InlineShapes.Count > 0 Or para.Range.ShapeRange.Count > 0 Then
            
            Dim shp As Object
            Dim imgWidth As Double
            Dim imgHeight As Double
            
            If para.Range.InlineShapes.Count > 0 Then
                Set shp = para.Range.InlineShapes(1)
            Else
                Set shp = para.Range.ShapeRange(1)
            End If
            
            ' Get image dimensions
            imgWidth = shp.Width
            imgHeight = shp.Height
            
            cellWidthInPoints = excelSheet.Cells(startingRowIndex, 2).Width
            cellHeightInPoints = excelSheet.Cells(startingRowIndex, 2).Height
            
            Cell_Horiz_Cnt = WorksheetFunction.RoundUp(imgWidth / cellWidthInPoints, 0)
            Cell_Vert_Cnt = WorksheetFunction.RoundUp(imgHeight / cellHeightInPoints, 0)
            If Cell_Horiz_Cnt > maxColumnCount Then
                Cell_Horiz_Cnt = maxColumnCount
            End If
            
            On Error GoTo ImageError
            
            ' Try different methods to copy the shape
            Dim copyMethod As Integer
            For copyMethod = 1 To 3
                On Error Resume Next
                Select Case copyMethod
                    Case 1
                        shp.Copy
                    Case 2
                        para.Range.Copy
                    Case 3
                        shp.Select
                        wordApp.Selection.Copy
                End Select
                If Err.Number = 0 Then Exit For
                On Error GoTo 0
            Next copyMethod
            
            If Err.Number <> 0 Then GoTo ImageError
            
            ' Paste into Excel
            excelSheet.Paste Destination:=excelSheet.Cells(startingRowIndex, 2)
            
            ' Get the pasted image shape
            Set pastedShape = excelSheet.Shapes(excelSheet.Shapes.Count)
            With pastedShape
                .Top = excelSheet.Cells(startingRowIndex, 2).Top
                .Left = excelSheet.Cells(startingRowIndex, 2).Left
                .Width = cellWidthInPoints * Cell_Horiz_Cnt
                .Height = cellHeightInPoints * Cell_Vert_Cnt
                .Placement = xlMoveAndSize ' Ensure the image is moved and sized with cells
            End With
            
            ' Add borders
            excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount + 1), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
            
            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, 1))
                ' Clear any existing borders
                .Borders.LineStyle = xlNone
                
                ' Add outer borders only - without using With
                .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlAutomatic
            End With
            startingRowIndex = startingRowIndex + Cell_Vert_Cnt + 1
            
            GoTo ImageHandled
        
ImageError:
                ' An error occurred, so increment the error counter
                ErrorCounter_IMG = ErrorCounter_IMG + 1
                ' Reset the error
                Err.Clear

                With excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt - 1, startingRowIndex + Cell_Horiz_Cnt - 1))
                    If .MergeCells Then
                        .MergeCells = False
                    End If
                End With

                ' Write "Failed Table Space"
                excelSheet.Range(excelSheet.Cells(startingRowIndex, 2), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Merge
                excelSheet.Cells(startingRowIndex, 2).Value = "Placeholder - Failed To Copy Image"
                excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount + 4)).Interior.Color = RGB(255, 255, 0)

                ' Increment the starting row index
                startingRowIndex = startingRowIndex + 2
        
ImageHandled:

        Else
        
            ' Content for non-heading paragraphs
            ' Your existing code for writing content to Excel worksheet
            ' Increment the starting row index
            
            If excelSheet.Cells(startingRowIndex, 1).Value <> "" Then
                startingRowIndex = startingRowIndex + 1
            Else
                'Do nothing
            End If
            
            If tableHandled Then
                ' Move to the next paragraph
                Set para = para.Next
                ' Reset the flag
                tableHandled = False
            End If

            
            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)
            
            If Len(cleanedText) > 1 And (cleanedText <> "") Then
                ' Check if text starts with lowercase
                If (Left(cleanedText, 1) Like "[a-z]" And (para.Style = "Normal" Or para.Style = "Body Text")) Or (Left(cleanedText, 1) Like "[a-z]" And Not Right(excelSheet.Cells(startingRowIndex - 1, 1).Value, 1) = "." And (para.Style = "Normal" Or para.Style = "Body Text") And Not excelSheet.Cells(startingRowIndex - 1, 1).Font.Bold) Or ((para.Style = "Normal" Or para.Style = "Body Text") And Not Right(cleanedText, 1) = ":" And Not Left(cleanedText, 1) Like "[a-z]" And excelSheet.Cells(startingRowIndex - 1, 1).IndentLevel < 3) Then
                    ' Add to previous cell with space
                    excelSheet.Cells(startingRowIndex - 1, 1).Value = excelSheet.Cells(startingRowIndex - 1, 1).Value & " " & cleanedText
                    ' Reapply wrap and autofit
                    excelSheet.Rows(startingRowIndex - 1).WrapText = True
                    excelSheet.Rows(startingRowIndex - 1).EntireRow.AutoFit
                Else
            
                    If Len(cleanedText) > 1 And (cleanedText <> "") Then
                        
                        If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                            Tab_parts = Split(cleanedText, vbTab)
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                            Tab_parts = Split(cleanedText, vbTab)
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        Else
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = cleanedText
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        End If
                        
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                        End With
                        
                        If para.Range.Font.Underline <> wdUnderlineNone Then
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Underline = True
                        End If
        
                       ' Increment the starting row index
                        startingRowIndex = startingRowIndex + 1
        
                    End If
                End If
            End If
        End If
        
        paraIndex = paraIndex + 1
        
    Next para

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Record the end time
    endTime = Timer
    
    ' Calculate the duration
    duration = endTime - startTime

    If Group_Headings = True Then
        If Count > 1 Then
            ReDim Preserve Ending_Nums(0 To Count - 2)
            Ending_Nums(Count - 2) = startingRowIndex - 1
        
            ' Loop through each element in the arrays
            For i = LBound(Starting_Nums) To UBound(Starting_Nums)
                ' Group the rows in the Excel worksheet
                excelSheet.Rows(CInt(Starting_Nums(i)) & ":" & CInt(Ending_Nums(i))).Rows.Group
            Next i
        Else
            MsgBox ("Headings not picked up. Groupings will not be added.")
        End If
    End If
    
    
    If ErrorCounter_IMG > 0 Then
        MsgBox ("The amount of images that failed to copy: " & ErrorCounter_IMG & ". ")  '& " out of " & wordDoc.Tables.Count & ". "
    End If
    
    If ErrorCounter > 0 Then
        MsgBox ("The amount of tables that failed to copy: " & ErrorCounter & " out of " & wordDoc.Tables.Count & ". ")
    End If
    
    If ErrorCounter_TXTBOX > 0 Then
        MsgBox ("The amount of shapes that failed to copy: " & ErrorCounter_TXTBOX & ". ")
    End If
    
    
    MsgBox ("Total runtime: " & Int(duration / 60) & " minutes and " & WorksheetFunction.RoundUp((duration Mod 60), 0) & " seconds.")

    
    Application.EnableEvents = savedEnableEvents
    
    
    ' Close Word
    ' Save the document
    If docOpened = True Then
        
    Else
        wordDoc.Save
        wordDoc.Close
        wordApp.Quit
    End If
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    ThisWorkbook.Save
    
End Sub



Sub ImportWordDataToExcel_2()
    
    ' Set the Excel application and worksheet
    Dim excelApp As Object
    Set excelApp = ThisWorkbook.Application ' Assumes the code is in the Excel workbook
      
    
    Dim startTime As Double
    Dim endTime As Double
    Dim duration As Double
    
    ' Record the start time
    startTime = Timer
    
    Dim savedEnableEvents As Boolean
    savedEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim filePath As String
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    ' Configure the file dialog
    With fileDialog
        .InitialFileName = GetWorkbookPath()
        .AllowMultiSelect = False
        .Title = "Select a Word file"
        .Filters.Clear
        .Filters.Add "Word and Rich Text Files", "*.docx; *.doc; *.rtf"
    End With
    
    ' Show the file dialog
    If fileDialog.Show = -1 Then
        filePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "No file selected"
        Exit Sub
    End If
        
    Dim excelSheet As Object
    Dim FileName As String
    Dim MainExcelSheet As Object
    ' Extract document name from file path
    FileName = ExtractDocumentName(filePath)

    ' Set the source sheet to copy from
    Dim sourceSheet As Object
    Set sourceSheet = ThisWorkbook.Sheets("Reference_Sheet_Microsoft")
    Set MainExcelSheet = ThisWorkbook.Sheets("Main")
    ' Copy the source sheet
    sourceSheet.Copy After:=Sheets(Sheets.Count)
    ' Set the new sheet
    Set excelSheet = ActiveSheet
    ' Rename the new sheet
    If Len(FileName) > 31 Then
        excelSheet.Name = Left(FileName, 31)
    Else
        excelSheet.Name = FileName
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    Dim lastRow As Long
    lastRow = excelSheet.Cells(excelSheet.Rows.Count, "B").End(xlUp).Row

    If lastRow >= 9 Then
        
        ' Unhide rows within the range
        excelSheet.Rows("9:" & lastRow).EntireRow.Hidden = False

        ' Delete the range, but keep rows 10 and 11
        If lastRow > 9 Then
            excelSheet.Rows("10:" & lastRow).Delete
        End If
    End If

    excelSheet.Cells(1, 1).Value = FileName
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim objDoc As Object
    Dim docOpened As Boolean
    
    ' Check if there is an existing instance of Word open
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo 0
    
    If Not wordApp Is Nothing Then
        ' Check if the desired document is already open
        For Each objDoc In wordApp.Documents
            If objDoc.FullName = filePath Then
                ' Document is already open, use the existing instance
                Set wordDoc = objDoc
                docOpened = True
                Exit For
            End If
        Next objDoc
    Else
        Set wordApp = CreateObject("Word.Application")
    End If
    
    ' If the document is not open, open it
    If Not docOpened Then
        Set wordDoc = wordApp.Documents.Open(filePath)
    End If
    
    RemoveHeadersAndFooters wordDoc

    ' Set the starting row index
    Dim startingRowIndex As Integer
    startingRowIndex = 1
    
    Dim numLines As Integer
    Dim cleanedText As String
    Dim Starting_Nums() As Variant
    Dim Ending_Nums() As Variant
    Dim First_Heading As Boolean
    Dim Count As Integer
    Dim End_Count As Integer
    First_Heading = False
    Count = 1
    End_Count = 1
    Dim Appendix As Boolean
    Dim pastedTables As Object
    Dim pastedTables_Create As Object
    Set pastedTables = CreateObject("Scripting.Dictionary")
    Set pastedTables_Create = CreateObject("Scripting.Dictionary")
    Set PastedImage = CreateObject("Scripting.Dictionary")
    Set PastedImage_Inline = CreateObject("Scripting.Dictionary")
    Dim ErrorCounter As Integer
    Dim ErrorCounter_IMG As Integer
    Dim tableHandled As Boolean
    Dim currentPageNumber As Long
    Dim maxColumnCount As Integer
    maxColumnCount = 0
    Dim currentPageNumber_Create As Integer
    Dim textWidth As Double
    Dim lineHeight As Double
    Dim Max_Char_Cnt As Integer
    Dim App_A1_isTrue As Boolean
    Dim Tab_parts() As String
    Dim Heading_Size_Min As Integer
    Dim Heading_Size_Max As Integer
    Dim spacePos As Integer
    Dim tabPos As Integer
    Dim firstDelimiterPos As Integer
    Dim headingParts() As String
    Dim paraIndex As Integer
    Dim Table_Image As Boolean
    Dim Group_Headings As Boolean
    paraIndex = 1

    Dim detectedSizes As Variant
    detectedSizes = AutoDetectHeadingSizes(wordDoc)
    
    Heading_Size_Min = detectedSizes(1)
    Heading_Size_Max = detectedSizes(2)
    
    Table_Image = MainExcelSheet.Cells(5, 8).Value
    Group_Headings = MainExcelSheet.Cells(3, 8).Value
    
    If Table_Image = False Then
        MsgBox ("This template doesn't support copying Word tables. Tables will be copied as images.")
        Table_Image = True
    End If
    
    If Heading_Size_Min = 0 Then
        MsgBox "The textsize of the paragraph was not set, a size of 10 will be used."
        Heading_Size_Min = 10
    End If
    
    App_A1_isTrue = False
    
    lineHeight = 11 * 1.6

    maxColumnCount = 1

    'Where the first excel entry will be
    startingRowIndex = 10
    
    Max_Char_Cnt = 135
    
    Dim test_para As String
    
    For Each para In wordDoc.Paragraphs
    
        If para.Range.Information(3) <> currentPageNumber Then ' Assuming 3 corresponds to wdActiveEndPageNumber
            ' Update the page number
            currentPageNumber = para.Range.Information(3)
        End If
        
        
        test_para = para.Range.text
        
        Appendix = False
    
        If (Not para.Range.Tables.Count > 0) And ((para.Style Like "Heading*" And Len(TrimSpacesAndTabs(para.Range.text)) > 0) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And IsEntireTextBold(para.Range) And para.Range.Font.size >= Heading_Size_Min And para.Style <> "List Paragraph") Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextBold(para.Range) And para.Range.Font.size >= Heading_Size_Min) Or (para.Range.Font.Bold = False And para.Style = "List Paragraph" And (UBound(Split(para.Range.ListFormat.ListString, ".")) > 1) And para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size <= 100)) Or _
        (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextCapitals(para.Range) And IsNumeric(Left(para.Range.ListFormat.ListString, 1))) Or (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style <> "List Paragraph" And IsEntireTextCapitals(para.Range) And para.Range.Font.size >= Heading_Size_Min And para.Range.Tables.Count = 0) Then


            ' Extract heading number and text
            Dim headingNumber As String
            Dim Heading_Level As Integer
            Dim headingText As String
            
            If Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style Like "Heading*" Then
                headingNumber = para.Range.ListFormat.ListString
                Heading_Level = para.OutlineLevel
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Range.Font.size >= Heading_Size_Min And IsEntireTextBold(para.Range) And para.Style <> "List Paragraph" Then
                ' Find the position of the first space and the first tab
                spacePos = InStr(para.Range.text, " ")
                tabPos = InStr(para.Range.text, vbTab)
                ' Determine which comes first (ignoring zeroes)
                If spacePos > 0 And tabPos > 0 Then
                    firstDelimiterPos = WorksheetFunction.Min(spacePos, tabPos)
                ElseIf spacePos > 0 Then
                    firstDelimiterPos = spacePos
                ElseIf tabPos > 0 Then
                    firstDelimiterPos = tabPos
                Else
                    firstDelimiterPos = 0
                End If
                
                If firstDelimiterPos >= 1 Then
                    headingNumber = Left(para.Range.text, firstDelimiterPos - 1)
                Else
                    headingNumber = ""
                End If

                
                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")
                
                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And para.Range.Font.size >= Heading_Size_Min And IsEntireTextBold(para.Range) Then
                headingNumber = para.Range.ListFormat.ListString
                Heading_Level = para.Range.ListFormat.ListLevelNumber
            ElseIf (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Range.Font.Bold = False And para.Style = "List Paragraph" And (UBound(Split(para.Range.ListFormat.ListString, ".")) > 1) And para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size <= 100) Then
                headingNumber = para.Range.ListFormat.ListString

                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")

                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            ElseIf Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style = "List Paragraph" And IsEntireTextCapitals(para.Range) And IsNumeric(Left(para.Range.ListFormat.ListString, 1)) Then
                headingNumber = para.Range.ListFormat.ListString

                ' Use Split to differentiate between second and third-level headings
                headingParts = Split(headingNumber, ".")

                If (IsNumeric(headingNumber) And UBound(headingParts) = 0) Or ((headingNumber Like "*\.0") And UBound(headingParts) = 0) Then ' And para.Range.Font.Size >= Heading_Size_Min
                    Heading_Level = 1
                ElseIf UBound(headingParts) = 1 Then
                    Heading_Level = 2
                ElseIf UBound(headingParts) = 2 Then
                    Heading_Level = 3
                ElseIf UBound(headingParts) = 3 Then
                    Heading_Level = 4
                ElseIf para.Range.Font.size >= Heading_Size_Min And para.Range.Font.size < Heading_Size_Max Then
                    Heading_Level = 2
                ElseIf para.Range.Font.size >= Heading_Size_Max Then
                    Heading_Level = 1
                Else
                    Heading_Level = 2
                End If
            ElseIf (Len(TrimSpacesAndTabs(para.Range.text)) > 0 And para.Style <> "List Paragraph" And IsEntireTextCapitals(para.Range) And para.Range.Font.size >= Heading_Size_Min) And para.Range.Tables.Count = 0 Then
                Heading_Level = 1
            End If
            
            headingNumber = para.Range.ListFormat.ListString
            headingText = para.Range.text
            
            If tableHandled Then
                ' Move to the next paragraph
                Set para = para.Next
                ' Reset the flag
                tableHandled = False
            End If
            
            

            ' Check if it's the highest level heading
            If Len(headingNumber) > 0 Then
                If Heading_Level = 1 Then
                    
                    ReDim Preserve Starting_Nums(0 To Count - 1)
                    Starting_Nums(Count - 1) = startingRowIndex + 1
                    
                    If First_Heading = True Then
                        ReDim Preserve Ending_Nums(0 To Count - 2)
                        Ending_Nums(Count - 2) = startingRowIndex - 1
                        End_Count = End_Count + 1
                    End If
                    
                    Count = Count + 1
                    First_Heading = True
                    
                    
                    ' Merge, format, and add borders to the merged cell for top-level heading
                    With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                        .HorizontalAlignment = -4131 ' Left align
                        .VerticalAlignment = -4108 ' Center vertically
                        .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                        .Borders.LineStyle = xlContinuous ' Set outside borders
                    End With
    
                    ' Set formatting for top-level heading
                    With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                        .Name = "Times New Roman" ' Font name
                        .size = 10 ' Font size
                        .Bold = True ' Bold
                    End With
                    excelSheet.Cells(startingRowIndex, 1).IndentLevel = 0
                Else
                    If Heading_Level = 2 Then
                        ' Second-level heading
                        ' Your code for second-level formatting
                        ' Merge, format, and add borders to the merged cell for second-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
                        ' Set formatting for second-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                            .Name = "Times New Roman" ' Font name
                            .size = 10 ' Font size
                            .Bold = True ' Bold
                        End With
                        excelSheet.Cells(startingRowIndex, 1).IndentLevel = 1
                    ElseIf Heading_Level = 3 Then
                        ' Third-level heading
                        ' Your code for third-level formatting
                        ' Merge, format, and add borders to the merged cell for third-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
    
                        ' Set formatting for third-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                            .Name = "Times New Roman" ' Font name
                            .size = 10 ' Font size
                            .Bold = False ' Bold
                        End With
                        excelSheet.Cells(startingRowIndex, 1).IndentLevel = 2
                    ElseIf Heading_Level = 4 Then
                        ' Fourth-level heading
                        ' Your code for third-level formatting
                        ' Merge, format, and add borders to the merged cell for third-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
    
                        ' Set formatting for third-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                            .Name = "Times New Roman" ' Font name
                            .size = 10 ' Font size
                            .Bold = False ' Bold
                        End With
                        excelSheet.Cells(startingRowIndex, 1).IndentLevel = 3
                    Else
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
        
                        ' Set formatting for top-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                            .Name = "Times New Roman" ' Font name
                            .size = 10 ' Font size
                            .Bold = False ' Bold
                        End With
                        excelSheet.Cells(startingRowIndex, 1).IndentLevel = 0
                    End If
                End If
                
            Else
                
                Dim appendixParts() As String
                
                If ((InStr(1, headingText, "APPENDIX") > 0) Or (InStr(1, headingText, "Appendix") > 0)) And ((InStr(1, headingText, "Publications Referenced")) = 0) Then
                   Appendix = True
                   
                   ' Split the heading into two parts based on the tab character
                    appendixParts = Split(headingText, vbTab)
                   ' Merge, format, and add borders to the merged cell for top-level heading
                   With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                       .HorizontalAlignment = -4131 ' Left align
                       .VerticalAlignment = -4108 ' Center vertically
                       .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                       .Borders.LineStyle = xlContinuous ' Set outside borders
                   End With
    
                   ' Set formatting for top-level heading
                   With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                       .Name = "Times New Roman" ' Font name
                       .size = 10 ' Font size
                       .Bold = True ' Bold
                   End With


                   ReDim Preserve Starting_Nums(0 To Count - 1)
                   Starting_Nums(Count - 1) = startingRowIndex + 1

                   ReDim Preserve Ending_Nums(0 To Count - 2)
                   Ending_Nums(Count - 2) = startingRowIndex - 1
                   End_Count = End_Count + 1
                   Count = Count + 1
                ElseIf (InStr(1, headingText, "Table") > 0) Or (InStr(1, headingNumber, "Table") > 0) Then
                
                    With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                        .HorizontalAlignment = -4131 ' Left align
                        .VerticalAlignment = -4108 ' Center vertically
                    End With
                    
                    ' Set formatting for third-level heading
                    With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                        .Name = "Times New Roman" ' Font name
                        .size = 10 ' Font size
                        .Bold = True ' Bold
                    End With
                    
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                    excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                    
                ElseIf ((InStr(1, headingText, "APPENDIX") > 0) And (InStr(1, headingText, "Publications Referenced")) > 0) Then
                
                    App_A1_isTrue = True
                    Appendix = True
                   
                   ' Split the heading into two parts based on the tab character
                    appendixParts = Split(headingText, vbTab)
                   ' Merge, format, and add borders to the merged cell for top-level heading
                   With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                       .HorizontalAlignment = -4131 ' Left align
                       .VerticalAlignment = -4108 ' Center vertically
                       .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                       .Borders.LineStyle = xlContinuous ' Set outside borders
                   End With
    
                   ' Set formatting for top-level heading
                   With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                        .Name = "Times New Roman" ' Font name
                        .size = 10 ' Font size
                        .Bold = True ' Bold
                   End With


                   ReDim Preserve Starting_Nums(0 To Count - 1)
                   Starting_Nums(Count - 1) = startingRowIndex + 1

                   ReDim Preserve Ending_Nums(0 To Count - 2)
                   Ending_Nums(Count - 2) = startingRowIndex - 1
                   End_Count = End_Count + 1

                   Count = Count + 1
                Else
                    If Heading_Level = 1 Then
                        
                        ReDim Preserve Starting_Nums(0 To Count - 1)
                        Starting_Nums(Count - 1) = startingRowIndex + 1
                        
                        If First_Heading = True Then
                            ReDim Preserve Ending_Nums(0 To Count - 2)
                            Ending_Nums(Count - 2) = startingRowIndex - 1
                            End_Count = End_Count + 1
                        End If
                        
                        Count = Count + 1
                        First_Heading = True
                        
                        
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                            .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                            .Borders.LineStyle = xlContinuous ' Set outside borders
                        End With
        
                        ' Set formatting for top-level heading
                        With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                            .Name = "Times New Roman" ' Font name
                            .size = 10 ' Font size
                            .Bold = True ' Bold
                        End With
                        excelSheet.Cells(startingRowIndex, 1).IndentLevel = 0
                    Else
        
                        If Heading_Level = 2 Then
                            ' Second-level heading
                            ' Your code for second-level formatting
                            ' Merge, format, and add borders to the merged cell for second-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
                            ' Set formatting for second-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                                .Name = "Times New Roman" ' Font name
                                .size = 10 ' Font size
                                .Bold = True ' Bold
                            End With
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 1
                        ElseIf Heading_Level = 3 Then
                            ' Third-level heading
                            ' Your code for third-level formatting
                            ' Merge, format, and add borders to the merged cell for third-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
        
                            ' Set formatting for third-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                                .Name = "Times New Roman" ' Font name
                                .size = 10 ' Font size
                                .Bold = True ' Bold
                            End With
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 2
                        ElseIf Heading_Level = 4 Then
                            ' Fourth-level heading
                            ' Your code for third-level formatting
                            ' Merge, format, and add borders to the merged cell for third-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
        
                            ' Set formatting for third-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                                .Name = "Times New Roman" ' Font name
                                .size = 10 ' Font size
                                .Bold = True ' Bold
                            End With
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 3
                        Else
                            ' Merge, format, and add borders to the merged cell for top-level heading
                            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                                .HorizontalAlignment = -4131 ' Left align
                                .VerticalAlignment = -4108 ' Center vertically
                                .Interior.Color = RGB(217, 217, 217) ' Color for the merged cell
                                .Borders.LineStyle = xlContinuous ' Set outside borders
                            End With
            
                            ' Set formatting for top-level heading
                            With excelSheet.Cells(startingRowIndex, 1).Resize(, maxColumnCount + 2).Font
                                .Name = "Times New Roman" ' Font name
                                .size = 10 ' Font size
                                .Bold = True ' Bold
                            End With
                        End If
                    End If
                
                End If
            End If

            If TrimSpacesAndTabs(headingText) <> "" Then
                ' Set the value for the merged cell
                If Appendix = False Then
                    If headingNumber <> "" Then
                        excelSheet.Cells(startingRowIndex, 1).Value = headingNumber & Space(1) & TrimSpacesAndTabs(headingText)
                    Else
                        If Len(TrimSpacesAndTabs(headingNumber)) > 0 Then
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(headingNumber) & Space(1) & TrimSpacesAndTabs(headingText)
                        Else
                            excelSheet.Cells(startingRowIndex, 1).Value = Space(1) & TrimSpacesAndTabs(headingText)
                        End If
                    End If
                Else
                    If UBound(appendixParts) > 0 Then
                        excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(appendixParts(0)) & Space(1) & TrimSpacesAndTabs(appendixParts(1))
                    Else
                        excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(appendixParts(0))
                    End If
                End If
            End If
            
            ' Increment the starting row index
            startingRowIndex = startingRowIndex + 1
                

        ElseIf para.Range.Tables.Count > 0 Then
            
            Set tbl = para.Range.Tables(1)
            Dim tableKey As String
            tableKey = para.Range.Tables(1).Range.Start & ":" & para.Range.Tables(1).Range.End & ":" & currentPageNumber
            
            If Not pastedTables.Exists(tableKey) Then
                If Table_Image Then
                    ' Handle table as image
                    Dim tblWidth As Double
                    Dim tblHeight As Double
                    
                    
                    ' Copy again
                    tbl.Range.CopyAsPicture
                    Application.Wait Now + TimeValue("00:00:01")
                    DoEvents
                    
                    On Error Resume Next
                    
                    excelSheet.Cells(startingRowIndex, 1).PasteSpecial Paste:=xlPasteEnhancedMetafile
                    
                    If Err.Number <> 0 Then
                        
                        Err.Clear
                        ErrorCounter = ErrorCounter + 1
                        'GoTo RegularTablePasting Disabled because column count hasn't been adjusted to accomodate tables that have more columns than what is available on the template.
                    End If
                    
                    On Error GoTo 0
                    
                    ' Get the pasted image shape
                    Dim pastedShape As Shape
                    Set pastedShape = excelSheet.Shapes(excelSheet.Shapes.Count)
                    
                    If Not pastedShape Is Nothing Then
                        tblWidth = pastedShape.Width
                        tblHeight = pastedShape.Height
                        
                        Dim cellWidthInPoints As Double
                        Dim cellHeightInPoints As Double
                        cellWidthInPoints = excelSheet.Cells(startingRowIndex, 1).Width
                        cellHeightInPoints = excelSheet.Cells(startingRowIndex, 1).Height
                        
                        Dim Cell_Horiz_Cnt As Integer
                        Dim Cell_Vert_Cnt As Integer
                        Cell_Horiz_Cnt = WorksheetFunction.RoundUp(tblWidth / cellWidthInPoints, 0)
                        Cell_Vert_Cnt = WorksheetFunction.RoundUp(tblHeight / cellHeightInPoints, 0)
                        
                        Dim ScaleFactor As Double
                        ScaleFactor = 1
                        If Cell_Horiz_Cnt > maxColumnCount Then
                            ScaleFactor = (Cell_Horiz_Cnt / maxColumnCount) - 1
                            Cell_Horiz_Cnt = maxColumnCount
                            
                            pastedShape.Height = cellHeightInPoints * Cell_Vert_Cnt * ScaleFactor
                            pastedShape.Width = cellWidthInPoints * Cell_Horiz_Cnt * ScaleFactor
                            pastedShape.Top = excelSheet.Cells(startingRowIndex, 1).Top
                            pastedShape.Left = excelSheet.Cells(startingRowIndex, 1).Left
                        End If
                        
                        pastedShape.Placement = xlMoveAndSize ' Ensure the image is moved and sized with cells
                            
                        ' Add borders
                        excelSheet.Range(excelSheet.Cells(startingRowIndex, maxColumnCount + 1), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                        
                        
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex + Cell_Vert_Cnt + 1, 1))
                            ' Clear any existing borders
                            .Borders.LineStyle = xlNone
                            
                            ' Add outer borders only - without using With
                            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlAutomatic
                        End With
                        
                        ' Update the starting row index
                        startingRowIndex = startingRowIndex + Cell_Vert_Cnt + 1
                    End If
                End If
                ' Mark the table as pasted
                pastedTables.Add tableKey, True
            End If
            
            On Error GoTo 0
            
        ElseIf para.Range.ListFormat.ListType <> wdListNoNumbering Then
            
            Dim listLevel As Integer
            Dim Bullet_Char As String
            
            listLevel = para.Range.ListFormat.ListLevelNumber
            
            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)
            
            If listLevel <= 1 Then
                Bullet_Char = ChrW(&H2022)
            ElseIf (listLevel > 1) And (listLevel <= 2) Then
                Bullet_Char = ChrW(&H25CB)
            ElseIf (listLevel > 2) And (listLevel <= 3) Then
                Bullet_Char = ChrW(&H25C7)
            Else
                Bullet_Char = para.Range.ListFormat.ListString
            End If
                        
            Dim firstChar As String
            firstChar = Left(para.Range.ListFormat.ListString, 1)
            
            
            If para.Range.ListFormat.ListType = wdListSimpleNumbering Or para.Range.ListFormat.ListType = wdListOutlineNumbering Then
                
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = para.Range.ListFormat.ListString & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = para.Range.ListFormat.ListString & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Cells(startingRowIndex, 1).Value = para.Range.ListFormat.ListString & " " & cleanedText
                End If
                excelSheet.Cells(startingRowIndex, 1).IndentLevel = listLevel + 3
            
            ElseIf IsNumeric(firstChar) And Mid(para.Range.ListFormat.ListString, 3, 1) <> "." Then 'And Mid(para.Range.ListFormat.ListString, 3, 1) <> "."
                'Paste just the content in Excel
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Cells(startingRowIndex, 1).Value = cleanedText
                End If
                   
            
            ElseIf para.Range.ListFormat.ListType = wdListBullet Then
            
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                    'excelSheet.Cells(startingRowIndex, 1).Value = para.Range.ListFormat.ListString & " " & cleanedText
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & cleanedText
                End If
                excelSheet.Cells(startingRowIndex, 1).IndentLevel = listLevel + 3
            Else
                
                If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                    Tab_parts = Split(cleanedText, vbTab)
                    ' Paste the paragraph in the merged cell
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                Else
                    excelSheet.Cells(startingRowIndex, 1).Value = Bullet_Char & " " & cleanedText
                End If

                excelSheet.Cells(startingRowIndex, 1).IndentLevel = listLevel
            End If
            
            excelSheet.Rows(startingRowIndex).WrapText = True
            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount)).Font.Name = "Times New Roman"
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount)).Font.size = 10
            
            If para.Range.Font.Underline <> wdUnderlineNone Then
                excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Underline = True
            End If
            
            ' Merge, format, and add borders to the merged cell for top-level heading
            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount))
                .HorizontalAlignment = -4131 ' Left align
                .VerticalAlignment = -4108 ' Center vertically
            End With

            
            startingRowIndex = startingRowIndex + 1
            
        ElseIf para.Style = "Table Caption" Then

            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)

            excelSheet.Cells(startingRowIndex, 1).Value = cleanedText
            excelSheet.Rows(startingRowIndex).WrapText = True
            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount)).Font.Name = "Times New Roman"
            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount)).Font.size = 10
            
            ' Merge, format, and add borders to the merged cell for top-level heading
            With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                .HorizontalAlignment = -4131 ' Left align
                .VerticalAlignment = -4108 ' Center vertically
            End With
            
            ' Calculate the number of lines in the paragraph
            startingRowIndex = startingRowIndex + 1
            
        Else
        
            ' Content for non-heading paragraphs
            ' Your existing code for writing content to Excel worksheet
            ' Increment the starting row index
            
            If excelSheet.Cells(startingRowIndex, 1).Value <> "" Then
                startingRowIndex = startingRowIndex + 1
            Else
                'Do nothing
            End If
            
            If tableHandled Then
                ' Move to the next paragraph
                Set para = para.Next
                ' Reset the flag
                tableHandled = False
            End If

            
            cleanedText = Trim(para.Range.text)
            cleanedText = RemoveNumbering(cleanedText)
            cleanedText = TrimSpacesAndTabs(cleanedText)

            If Len(cleanedText) > 1 And (cleanedText <> "") Then
                ' Check if text starts with lowercase
                If (Left(cleanedText, 1) Like "[a-z]" And (para.Style = "Normal" Or para.Style = "Body Text")) Or (Left(cleanedText, 1) Like "[a-z]" And Not Right(excelSheet.Cells(startingRowIndex - 1, 1).Value, 1) = "." And (para.Style = "Normal" Or para.Style = "Body Text") And Not excelSheet.Cells(startingRowIndex - 1, 1).Font.Bold) Or ((para.Style = "Normal" Or para.Style = "Body Text") And Not Right(cleanedText, 1) = ":" And Not Left(cleanedText, 1) Like "[a-z]" And excelSheet.Cells(startingRowIndex - 1, 1).IndentLevel > 3) Then
                    
                    ' Add to previous cell with space
                    excelSheet.Cells(startingRowIndex - 1, 1).Value = excelSheet.Cells(startingRowIndex - 1, 1).Value & " " & cleanedText
                    ' Reapply wrap and autofit
                    excelSheet.Rows(startingRowIndex - 1).WrapText = True
                    excelSheet.Rows(startingRowIndex - 1).EntireRow.AutoFit
                Else
            
                    If Len(cleanedText) > 1 And (cleanedText <> "") Then
                        
                        If (App_A1_isTrue = True) And (InStr(cleanedText, vbTab) > 0) Then
                            Tab_parts = Split(cleanedText, vbTab)
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        ElseIf (App_A1_isTrue = False) And (InStr(cleanedText, vbTab) > 0) Then
                            Tab_parts = Split(cleanedText, vbTab)
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = TrimSpacesAndTabs(Tab_parts(0)) & Space(1) & TrimSpacesAndTabs(Tab_parts(1))
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        Else
                            ' Paste the paragraph in the merged cell
                            excelSheet.Cells(startingRowIndex, 1).Value = cleanedText
                            ' Wrap text in the merged cell
                            excelSheet.Rows(startingRowIndex).WrapText = True
                            excelSheet.Rows(startingRowIndex).EntireRow.AutoFit
                            ' Put borders around all cells in the line from A to M
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Borders.LineStyle = xlContinuous
                            ' Set font to Times New Roman, size 11
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.Name = "Times New Roman"
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2)).Font.size = 10
                            excelSheet.Cells(startingRowIndex, 1).IndentLevel = 4
                        End If
                        
                        ' Merge, format, and add borders to the merged cell for top-level heading
                        With excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, maxColumnCount + 2))
                            .HorizontalAlignment = -4131 ' Left align
                            .VerticalAlignment = -4108 ' Center vertically
                        End With
                        
                        If para.Range.Font.Underline <> wdUnderlineNone Then
                            excelSheet.Range(excelSheet.Cells(startingRowIndex, 1), excelSheet.Cells(startingRowIndex, 1 + maxColumnCount)).Font.Underline = True
                        End If
        
                       ' Increment the starting row index
                        startingRowIndex = startingRowIndex + 1
        
                    End If
                End If
            End If
        End If
        
        paraIndex = paraIndex + 1
        
    Next para

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Record the end time
    endTime = Timer
    
    ' Calculate the duration
    duration = endTime - startTime

    If Group_Headings = True Then
        If Count > 1 Then
            ReDim Preserve Ending_Nums(0 To Count - 2)
            Ending_Nums(Count - 2) = startingRowIndex - 1
        
            ' Loop through each element in the arrays
            For i = LBound(Starting_Nums) To UBound(Starting_Nums)
                ' Group the rows in the Excel worksheet
                excelSheet.Rows(CInt(Starting_Nums(i)) & ":" & CInt(Ending_Nums(i))).Rows.Group
            Next i
        Else
            MsgBox ("Headings not picked up. Groupings will not be added.")
        End If
    End If
    
    If ErrorCounter_IMG > 0 Then
        MsgBox ("The amount of images that failed to copy: " & ErrorCounter_IMG & ". ")  '& " out of " & wordDoc.Tables.Count & ". "
    End If
    
    If ErrorCounter > 0 Then
        MsgBox ("The amount of tables that failed to copy: " & ErrorCounter & " out of " & wordDoc.Tables.Count & ". ")
    End If
    
    If ErrorCounter_TXTBOX > 0 Then
        MsgBox ("The amount of shapes that failed to copy: " & ErrorCounter_TXTBOX & ". ")
    End If
    
    
    MsgBox ("Total runtime: " & Int(duration / 60) & " minutes and " & WorksheetFunction.RoundUp((duration Mod 60), 0) & " seconds.")

    
    Application.EnableEvents = savedEnableEvents
    
    
    ' Close Word
    ' Save the document
    If docOpened = True Then
        
    Else
        wordDoc.Save
        wordDoc.Close
        wordApp.Quit
    End If
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    ThisWorkbook.Save
    
End Sub



Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.Path
    If Right(GetWorkbookPath, 1) <> "\" Then
        GetWorkbookPath = GetWorkbookPath & "\"
    End If
End Function

Function RemoveNumbering(text As String) As String
    ' Remove leading digits, tabs, and whitespace
    Dim Placeholder As String
    
    Placeholder = text
    
    'If (Len(Placeholder) > 1) And ((Left(Placeholder, 1) Like "[0-9]") And (Mid(Placeholder, 2, 1) Like "[0-9]")) And ((Mid(Placeholder, 3, 1) = " ") Or (Mid(Placeholder, 3, 1) = vbTab)) Then
    If (Len(Placeholder) > 2) And ((Left(Placeholder, 1) Like "[0-9]") And (Mid(Placeholder, 2, 1) Like "[0-9]") And (Mid(Placeholder, 3, 1) = " " Or Mid(Placeholder, 3, 1) = vbTab)) Then
        Do While Len(Placeholder) > 1 And (Left(Placeholder, 1) Like "[0-9]" Or Left(Placeholder, 1) = vbTab Or Left(Placeholder, 1) = " ")
            Placeholder = Right(Placeholder, Len(Placeholder) - 1)
        Loop
    ElseIf (Len(Placeholder) <= 3) And (Left(Placeholder, 1) Like "[0-9]") And (Mid(Placeholder, 2, 1) Like "[0-9]") Then
        Placeholder = ""
    End If
    ' Trim the text
    RemoveNumbering = Trim(Placeholder)
End Function


Function TrimSpacesAndTabs(inputText As String) As String
    ' Check for empty or null input
    If inputText = "" Or IsNull(inputText) Then
        TrimSpacesAndTabs = ""
        Exit Function
    End If
    
    Dim result As String
    result = inputText
    
    ' Remove common Word special characters
    result = Replace(result, vbCr, "")      ' Carriage return
    result = Replace(result, vbLf, "")      ' Line feed
    result = Replace(result, vbCrLf, "")    ' Windows line ending
    result = Replace(result, vbVerticalTab, "") ' Vertical tab
    result = Replace(result, vbFormFeed, "") ' Form feed
    result = Replace(result, Chr(7), "")     ' Bell
    result = Replace(result, Chr(160), " ")  ' Non-breaking space
    
    ' Handle multiple spaces and tabs
    Do While InStr(result, "  ") > 0    ' Double spaces
        result = Replace(result, "  ", " ")
    Loop
    
    Do While InStr(result, vbTab & vbTab) > 0    ' Double tabs
        result = Replace(result, vbTab & vbTab, vbTab)
    Loop
    
    ' Remove leading and trailing spaces and tabs
    result = Trim(result)
    
    ' Remove any remaining trailing special characters only if result is not empty
'    If Len(result) > 0 Then
'        Do While Len(result) > 0 And (AscW(Right(result, 1)) < 32 Or AscW(Right(result, 1)) = 160)
'            result = Left(result, Len(result) - 1)
'        Loop
'    End If

    If Len(result) > 0 Then
        On Error Resume Next
        Dim lastChar As String
        lastChar = Right(result, 1)
        If Err.Number = 0 Then  ' Only proceed if we successfully got the last character
            Do While Len(result) > 0 And (AscW(lastChar) < 32 Or AscW(lastChar) = 160)
                result = Left(result, Len(result) - 1)
                If Len(result) > 0 Then
                    lastChar = Right(result, 1)
                Else
                    Exit Do
                End If
            Loop
        End If
        On Error GoTo 0
    End If
    
    TrimSpacesAndTabs = result
End Function

Sub RemoveHeadersAndFooters(doc As Object)
    Dim sec As Object
    
    ' Loop through each section in the document
    For Each sec In doc.Sections
        ' Remove the header
        sec.Headers(wdHeaderFooterPrimary).Range.Delete
        sec.Headers(wdHeaderFooterFirstPage).Range.Delete
        sec.Headers(wdHeaderFooterEvenPages).Range.Delete
        
        ' Remove the footer
        sec.Footers(wdHeaderFooterPrimary).Range.Delete
        sec.Footers(wdHeaderFooterFirstPage).Range.Delete
        sec.Footers(wdHeaderFooterEvenPages).Range.Delete
    Next sec
End Sub

Function ExtractDocumentName(filePath As String) As String
    Dim fileNameWithExtension As String
    Dim fileNameWithoutExtension As String
    Dim lastBackslashIndex As Integer
    
    ' Find the index of the last backslash in the file path
    lastBackslashIndex = InStrRev(filePath, "\")
    
    ' Extract the file name with extension
    fileNameWithExtension = Mid(filePath, lastBackslashIndex + 1)
    
    ' Remove the file extension from the file name
    fileNameWithoutExtension = Left(fileNameWithExtension, InStrRev(fileNameWithExtension, ".") - 1)
    
    ' Return the document name
    ExtractDocumentName = fileNameWithoutExtension
End Function

Private Function IsEntireTextBold(rng As Object) As Boolean
    Dim char As Object
    Dim textLength As Long
    
    IsEntireTextBold = True
    
    ' Check if range is empty or contains special characters only
    If rng.Characters.Count = 0 Or Trim(rng.text) = "" Then
        IsEntireTextBold = False
        Exit Function
    End If
    
    ' Loop through each character in the range
    For Each char In rng.Characters
        ' Skip if it's a special character (like paragraph mark)
        If Len(Trim(char.text)) > 0 Then
            If Not char.Font.Bold Then
                IsEntireTextBold = False
                Exit Function
            End If
        End If
    Next char
End Function

Private Function IsEntireTextCapitals(rng As Object) As Boolean
    Dim char As Object
    Dim i As Integer
    
    i = 1
    IsEntireTextCapitals = True
    
    If rng.Characters.Count = 0 Or Trim(rng.text) = "" Then
        IsEntireTextCapitals = False
        Exit Function
    End If
    
    For Each char In rng.Characters
        If i >= 25 Then
            Exit Function
        Else
            Dim charCode As Integer
            charCode = AscW(char.text)
            
            ' Skip whitespace, carriage return, and common punctuation
            If charCode > 32 And charCode <> 13 Then
                If Not ((charCode >= 65 And charCode <= 90) Or (charCode >= 48 And charCode <= 57) Or charCode = 45 Or charCode = 46 Or charCode = 40 Or charCode = 41 Or charCode = 44 Or charCode = 38 Or charCode = 47 Or charCode = 58) Then
                    IsEntireTextCapitals = False
                    Exit Function
                End If
            End If
        End If
        i = i + 1
    Next char
End Function



Function AutoDetectHeadingSizes(wordDoc As Object) As Variant
    Dim fontSizes As Object
    Set fontSizes = CreateObject("Scripting.Dictionary")
    Dim result(1 To 2) As Integer
    Dim para As Object
    Dim size As Single
    Dim i As Integer
    
    ' Scan first 100 paragraphs (or all if document is shorter)
    For i = 1 To WorksheetFunction.Min(250, wordDoc.Paragraphs.Count)
        Set para = wordDoc.Paragraphs(i)
        
        ' Only count if it looks like real content
        If Len(TrimSpacesAndTabs(para.Range.text)) > 3 And _
           Not para.Range.text Like "*Page #*" And _
           Not para.Range.Information(wdWithInTable) Then
            
            size = para.Range.Font.size
            
            ' Round to nearest integer to group similar sizes
            size = Round(size)
            
            ' Only track reasonable sizes (8-30pt)
            If size >= 8 And size <= 30 Then
                If fontSizes.Exists(size) Then
                    fontSizes(size) = fontSizes(size) + 1
                Else
                    fontSizes.Add size, 1
                End If
            End If
        End If
    Next i
    
    ' If we found no text, use defaults
    If fontSizes.Count = 0 Then
        result(1) = 10  ' Default min
        result(2) = 14  ' Default max
        AutoDetectHeadingSizes = result
        Exit Function
    End If
    
    ' Find most common size (likely body text)
    Dim bodySize As Integer
    bodySize = GetMostFrequentSize(fontSizes)
    
    ' Find largest size above body (likely biggest heading)
    Dim maxSize As Integer
    maxSize = GetLargestSizeAbove(fontSizes, bodySize)
    
    ' Apply safeguards
    ' Minimum heading size should be at least body text size
    result(1) = bodySize
    If result(1) < 8 Then result(1) = 8   ' Never go below 10
    If result(1) > 16 Then result(1) = 16  ' If body is too large, default to 10
    If result(2) <= result(1) Then result(2) = 14
    
    ' Maximum heading size
    result(2) = maxSize
    'If result(2) < result(1) + 2 Then result(2) = result(1) + 4  ' Ensure reasonable gap
    If result(2) > 20 Then result(2) = 16  ' Cap at reasonable size
    
    AutoDetectHeadingSizes = result
End Function

Function GetMostFrequentSize(fontSizes As Object) As Integer
    Dim size As Variant
    Dim maxCount As Integer
    Dim mostFrequent As Integer
    
    maxCount = 0
    mostFrequent = 11  ' Default fallback
    
    ' Find size with highest count
    For Each size In fontSizes.Keys
        If fontSizes(size) > maxCount Then
            maxCount = fontSizes(size)
            mostFrequent = CInt(size)
        End If
    Next size
    
    ' Additional logic: if multiple sizes are close in frequency,
    ' prefer the smaller one (likely body text)
    For Each size In fontSizes.Keys
        If fontSizes(size) >= maxCount * 0.8 And CInt(size) < mostFrequent Then
            mostFrequent = CInt(size)
        End If
    Next size
    
    GetMostFrequentSize = mostFrequent
End Function

Function GetLargestSizeAbove(fontSizes As Object, bodySize As Integer) As Integer
    Dim size As Variant
    Dim largestSize As Integer
    Dim CountSize As Integer
    largestSize = bodySize  ' Start with body size
    
    ' Find largest size that appears at least twice (to avoid outliers)
    For Each size In fontSizes.Keys
        If CInt(size) > bodySize And CInt(size) > largestSize Then
            ' Require at least 2 occurrences to avoid title page outliers
            CountSize = fontSizes(size)
            If CountSize >= 2 Then
                largestSize = CInt(size)
            End If
        End If
    Next size
    
    ' If no larger size found, use bodySize + 4
    If largestSize = bodySize Then
        largestSize = bodySize + 4
    End If
    
    GetLargestSizeAbove = largestSize
End Function

