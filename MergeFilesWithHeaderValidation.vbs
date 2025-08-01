Sub MergeFilesWithHeaderValidation()
    Dim fileCount As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim selectedFiles() As String
    Dim fd As FileDialog
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    Dim headerRef As Variant
    Dim headerCurrent As Variant
    Dim allHeaders As Collection
    Dim missingHeaders As Collection
    Dim headerFound As Boolean
    Dim outputWs As Worksheet
    Dim lastRow As Long
    Dim fName As String
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim dataRowCount As Long, dataColCount As Long
    Dim userResponse As VbMsgBoxResult
    Dim missingHeadersText As String
    Dim headerValidationPassed As Boolean
    
    ' Ask user how many files they want to merge
    fileCount = Application.InputBox("How many files do you want to merge?", Type:=1)
    If fileCount < 1 Then
        MsgBox "Invalid file count. Operation cancelled."
        Exit Sub
    End If

    ReDim selectedFiles(1 To fileCount)

    ' File selection loop
    For i = 1 To fileCount
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select file " & i & " of " & fileCount
            .Filters.Clear
            .Filters.Add "Excel and CSV Files", "*.xls; *.xlsx; *.xlsm; *.csv"
            .AllowMultiSelect = False
            If .Show <> -1 Then
                MsgBox "File selection cancelled."
                Exit Sub
            End If
            selectedFiles(i) = .SelectedItems(1)
        End With
    Next i

    ' Disable screen flicker
    Application.ScreenUpdating = False
    
    ' First pass: Collect all unique headers from all files
    Set allHeaders = New Collection
    
    For i = 1 To fileCount
        fName = selectedFiles(i)
        
        ' Open file
        If LCase(Right(fName, 4)) = ".csv" Then
            Workbooks.OpenText Filename:=fName, DataType:=xlDelimited, Comma:=True
        Else
            Workbooks.Open Filename:=fName
        End If
        
        Set wbTemp = ActiveWorkbook
        Set wsTemp = wbTemp.Sheets(1)
        
        ' Get current header
        headerCurrent = Application.Transpose(Application.Transpose(wsTemp.Range("A1", wsTemp.Cells(1, wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column)).Value))
        
        ' Add headers to collection (avoiding duplicates)
        For j = 1 To UBound(headerCurrent)
            On Error Resume Next
            allHeaders.Add headerCurrent(j), CStr(headerCurrent(j))
            On Error GoTo 0
        Next j
        
        wbTemp.Close False
    Next i
    
    ' Second pass: Check each file for missing headers and prompt user to add them
    headerValidationPassed = False
    
    Do While Not headerValidationPassed
        headerValidationPassed = True
        
        For i = 1 To fileCount
            fName = selectedFiles(i)
            Set missingHeaders = New Collection
            
            ' Open file
            If LCase(Right(fName, 4)) = ".csv" Then
                Workbooks.OpenText Filename:=fName, DataType:=xlDelimited, Comma:=True
            Else
                Workbooks.Open Filename:=fName
            End If
            
            Set wbTemp = ActiveWorkbook
            Set wsTemp = wbTemp.Sheets(1)
            
            ' Get current header
            headerCurrent = Application.Transpose(Application.Transpose(wsTemp.Range("A1", wsTemp.Cells(1, wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column)).Value))
            
            ' Check for missing headers
            For j = 1 To allHeaders.Count
                headerFound = False
                For k = 1 To UBound(headerCurrent)
                    If CStr(allHeaders(j)) = CStr(headerCurrent(k)) Then
                        headerFound = True
                        Exit For
                    End If
                Next k
                
                If Not headerFound Then
                    missingHeaders.Add allHeaders(j)
                    headerValidationPassed = False
                End If
            Next j
            
            ' If missing headers found, inform user and ask to add them
            If missingHeaders.Count > 0 Then
                missingHeadersText = ""
                For j = 1 To missingHeaders.Count
                    missingHeadersText = missingHeadersText & "- " & missingHeaders(j) & vbCrLf
                Next j
                
                userResponse = MsgBox("The following headers are missing in file:" & vbCrLf & _
                                    fName & vbCrLf & vbCrLf & _
                                    "Missing headers:" & vbCrLf & missingHeadersText & vbCrLf & _
                                    "Would you like to add these missing headers to this file?" & vbCrLf & _
                                    "(Click Yes to add headers, No to cancel operation)", _
                                    vbYesNo + vbQuestion, "Missing Headers Detected")
                
                If userResponse = vbYes Then
                    ' Add missing headers to the file
                    Call AddMissingHeaders(wsTemp, missingHeaders)
                    wbTemp.Save
                    MsgBox "Missing headers have been added to the file. The headers will be re-validated.", vbInformation
                Else
                    wbTemp.Close False
                    Application.ScreenUpdating = True
                    MsgBox "Operation cancelled by user.", vbInformation
                    Exit Sub
                End If
            End If
            
            wbTemp.Close False
        Next i
        
        ' If we added headers, we need to refresh the allHeaders collection
        If Not headerValidationPassed Then
            Set allHeaders = New Collection
            
            For i = 1 To fileCount
                fName = selectedFiles(i)
                
                ' Open file
                If LCase(Right(fName, 4)) = ".csv" Then
                    Workbooks.OpenText Filename:=fName, DataType:=xlDelimited, Comma:=True
                Else
                    Workbooks.Open Filename:=fName
                End If
                
                Set wbTemp = ActiveWorkbook
                Set wsTemp = wbTemp.Sheets(1)
                
                ' Get current header
                headerCurrent = Application.Transpose(Application.Transpose(wsTemp.Range("A1", wsTemp.Cells(1, wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column)).Value))
                
                ' Add headers to collection (avoiding duplicates)
                For j = 1 To UBound(headerCurrent)
                    On Error Resume Next
                    allHeaders.Add headerCurrent(j), CStr(headerCurrent(j))
                    On Error GoTo 0
                Next j
                
                wbTemp.Close False
            Next i
        End If
    Loop
    
    ' Final validation: Ensure all files have exactly the same headers in the same order
    For i = 1 To fileCount
        fName = selectedFiles(i)
        
        ' Open file
        If LCase(Right(fName, 4)) = ".csv" Then
            Workbooks.OpenText Filename:=fName, DataType:=xlDelimited, Comma:=True
        Else
            Workbooks.Open Filename:=fName
        End If
        
        Set wbTemp = ActiveWorkbook
        Set wsTemp = wbTemp.Sheets(1)
        
        ' Get current header
        headerCurrent = Application.Transpose(Application.Transpose(wsTemp.Range("A1", wsTemp.Cells(1, wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column)).Value))
        
        ' Set reference header from first file
        If i = 1 Then
            headerRef = headerCurrent
        Else
            ' Compare headers (order and content)
            If UBound(headerRef) <> UBound(headerCurrent) Then
                MsgBox "Header count mismatch in file: " & vbCrLf & fName & vbCrLf & _
                       "Expected " & UBound(headerRef) & " headers, found " & UBound(headerCurrent), vbCritical
                wbTemp.Close False
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            For j = 1 To UBound(headerRef)
                If CStr(headerRef(j)) <> CStr(headerCurrent(j)) Then
                    MsgBox "Header order/content mismatch in file: " & vbCrLf & fName & vbCrLf & _
                           "Position " & j & ": Expected '" & headerRef(j) & "', Found '" & headerCurrent(j) & "'", vbCritical
                    wbTemp.Close False
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
            Next j
        End If
        
        wbTemp.Close False
    Next i

    ' Create new worksheet for merged data
    Set outputWs = ThisWorkbook.Sheets.Add
    outputWs.Name = "MergedOutput_" & Format(Now, "yyyymmdd_hhmmss")
    lastRow = 1

    ' Process each file for merging
    For i = 1 To fileCount
        fName = selectedFiles(i)
        
        ' Open file
        If LCase(Right(fName, 4)) = ".csv" Then
            Workbooks.OpenText Filename:=fName, DataType:=xlDelimited, Comma:=True
        Else
            Workbooks.Open Filename:=fName
        End If
        
        Set wbTemp = ActiveWorkbook
        Set wsTemp = wbTemp.Sheets(1)
        
        ' Write header to output (only for first file)
        If i = 1 Then
            headerCurrent = Application.Transpose(Application.Transpose(wsTemp.Range("A1", wsTemp.Cells(1, wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column)).Value))
            outputWs.Range("A1").Resize(1, UBound(headerCurrent)).Value = headerCurrent
            lastRow = 2
        End If
        
        ' Define data range (excluding header)
        dataRowCount = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        dataColCount = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
        
        If dataRowCount > 1 Then
            Set dataRange = wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(dataRowCount, dataColCount))
            dataArray = dataRange.Value
            outputWs.Cells(lastRow, 1).Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
            lastRow = lastRow + UBound(dataArray, 1)
        End If
        
        wbTemp.Close False
    Next i

    Application.ScreenUpdating = True
    MsgBox "Files merged successfully into sheet '" & outputWs.Name & "'." & vbCrLf & _
           "All headers were validated and matched before merging.", vbInformation
End Sub

' Helper subroutine to add missing headers to a worksheet
Sub AddMissingHeaders(ws As Worksheet, missingHeaders As Collection)
    Dim lastCol As Long
    Dim i As Integer
    
    ' Find the last column with data
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Add each missing header
    For i = 1 To missingHeaders.Count
        lastCol = lastCol + 1
        ws.Cells(1, lastCol).Value = missingHeaders(i)
        ' Add a placeholder value or leave empty - you can modify this as needed
        ' ws.Cells(2, lastCol).Value = "" ' Leave empty or add default value
    Next i
    
    ' Auto-fit columns for better visibility
    ws.Columns.AutoFit
End Sub