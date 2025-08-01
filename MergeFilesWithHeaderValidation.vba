Sub MergeFilesIfHeadersMatch()
    Dim fileCount As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim selectedFiles() As String
    Dim fd As FileDialog
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    Dim headerRef As Variant
    Dim headerCurrent As Variant
    Dim match As Boolean
    Dim outputWs As Worksheet
    Dim lastRow As Long
    Dim fName As String
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim dataRowCount As Long, dataColCount As Long
    Dim missingHeaders As String
    Dim extraHeaders As String
    Dim userResponse As VbMsgBoxResult
    Dim headerValidationPassed As Boolean
    Dim maxRetries As Integer
    Dim retryCount As Integer

    ' Ask user how many files they want to merge
    fileCount = Application.InputBox("How many files do you want to merge?", Type:=1)
    If fileCount < 1 Then
        MsgBox "Invalid file count. Operation cancelled."
        Exit Sub
    End If

    ReDim selectedFiles(1 To fileCount)
    maxRetries = 3 ' Maximum number of retry attempts for header validation

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

    ' Create new worksheet for merged data
    Set outputWs = ThisWorkbook.Sheets.Add
    outputWs.Name = "MergedOutput_" & Format(Now, "yyyymmdd_hhmmss")
    lastRow = 1

    ' Disable screen flicker
    Application.ScreenUpdating = False

    ' Header validation and processing loop
    For i = 1 To fileCount
        retryCount = 0
        headerValidationPassed = False
        
        Do While Not headerValidationPassed And retryCount < maxRetries
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
            
            ' Compare headers
            If i = 1 Then
                headerRef = headerCurrent
                ' Write header to output
                outputWs.Range("A1").Resize(1, UBound(headerRef)).Value = headerRef
                lastRow = 2
                headerValidationPassed = True
            Else
                ' Detailed header comparison
                missingHeaders = ""
                extraHeaders = ""
                
                ' Check for missing headers (present in reference but not in current)
                For j = 1 To UBound(headerRef)
                    match = False
                    For k = 1 To UBound(headerCurrent)
                        If Trim(UCase(headerRef(j))) = Trim(UCase(headerCurrent(k))) Then
                            match = True
                            Exit For
                        End If
                    Next k
                    If Not match Then
                        If missingHeaders <> "" Then missingHeaders = missingHeaders & ", "
                        missingHeaders = missingHeaders & "'" & headerRef(j) & "'"
                    End If
                Next j
                
                ' Check for extra headers (present in current but not in reference)
                For j = 1 To UBound(headerCurrent)
                    match = False
                    For k = 1 To UBound(headerRef)
                        If Trim(UCase(headerCurrent(j))) = Trim(UCase(headerRef(k))) Then
                            match = True
                            Exit For
                        End If
                    Next k
                    If Not match Then
                        If extraHeaders <> "" Then extraHeaders = extraHeaders & ", "
                        extraHeaders = extraHeaders & "'" & headerCurrent(j) & "'"
                    End If
                Next j
                
                ' Check if headers match exactly (same order and content)
                match = True
                If UBound(headerRef) <> UBound(headerCurrent) Then
                    match = False
                Else
                    For j = 1 To UBound(headerRef)
                        If Trim(UCase(headerRef(j))) <> Trim(UCase(headerCurrent(j))) Then
                            match = False
                            Exit For
                        End If
                    Next j
                End If
                
                If match Then
                    headerValidationPassed = True
                Else
                    ' Display detailed mismatch information
                    Dim errorMessage As String
                    errorMessage = "Header mismatch found in file:" & vbCrLf & vbCrLf
                    errorMessage = errorMessage & "File: " & Dir(fName) & vbCrLf & vbCrLf
                    
                    If missingHeaders <> "" Then
                        errorMessage = errorMessage & "Missing Variables/Headers:" & vbCrLf
                        errorMessage = errorMessage & missingHeaders & vbCrLf & vbCrLf
                    End If
                    
                    If extraHeaders <> "" Then
                        errorMessage = errorMessage & "Extra Variables/Headers:" & vbCrLf
                        errorMessage = errorMessage & extraHeaders & vbCrLf & vbCrLf
                    End If
                    
                    errorMessage = errorMessage & "Would you like to:" & vbCrLf
                    errorMessage = errorMessage & "• YES: Open the file to add/modify headers" & vbCrLf
                    errorMessage = errorMessage & "• NO: Cancel the merge operation" & vbCrLf
                    errorMessage = errorMessage & "• CANCEL: Skip this file and continue"
                    
                    Application.ScreenUpdating = True
                    userResponse = MsgBox(errorMessage, vbYesNoCancel + vbQuestion, "Header Validation Error")
                    Application.ScreenUpdating = False
                    
                    Select Case userResponse
                        Case vbYes
                            ' Keep the file open for user to edit
                            Application.ScreenUpdating = True
                            MsgBox "Please add the missing headers and/or remove extra headers in the currently open file." & vbCrLf & vbCrLf & _
                                   "Reference Headers (from first file):" & vbCrLf & Join(headerRef, ", ") & vbCrLf & vbCrLf & _
                                   "Current Headers:" & vbCrLf & Join(headerCurrent, ", ") & vbCrLf & vbCrLf & _
                                   "Click OK when you have finished editing the headers.", vbInformation, "Edit Headers"
                            Application.ScreenUpdating = False
                            
                            ' Don't close the workbook, let user edit it
                            retryCount = retryCount + 1
                            
                        Case vbNo
                            ' Cancel entire operation
                            wbTemp.Close False
                            Application.ScreenUpdating = True
                            outputWs.Delete
                            MsgBox "Merge operation cancelled by user.", vbInformation
                            Exit Sub
                            
                        Case vbCancel
                            ' Skip this file
                            wbTemp.Close False
                            headerValidationPassed = True ' Skip to next file
                            GoTo NextFile
                    End Select
                End If
            End If
            
            ' If validation failed and we're retrying, close and reopen the file
            If Not headerValidationPassed And retryCount < maxRetries Then
                wbTemp.Close True ' Save changes if any
            End If
            
        Loop
        
        ' If we exceeded max retries
        If Not headerValidationPassed And retryCount >= maxRetries Then
            Application.ScreenUpdating = True
            MsgBox "Maximum retry attempts exceeded for file: " & Dir(fName) & vbCrLf & _
                   "This file will be skipped.", vbExclamation
            Application.ScreenUpdating = False
            If Not wbTemp Is Nothing Then wbTemp.Close False
            GoTo NextFile
        End If
        
        ' Process data if headers are valid
        If headerValidationPassed And i > 1 Then ' i > 1 because first file is always processed
            ' Define data range (excluding header)
            dataRowCount = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
            dataColCount = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
            
            If dataRowCount > 1 Then
                Set dataRange = wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(dataRowCount, dataColCount))
                dataArray = dataRange.Value
                outputWs.Cells(lastRow, 1).Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
                lastRow = lastRow + UBound(dataArray, 1)
            End If
        ElseIf i = 1 Then
            ' Process first file data
            dataRowCount = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
            dataColCount = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
            
            If dataRowCount > 1 Then
                Set dataRange = wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(dataRowCount, dataColCount))
                dataArray = dataRange.Value
                outputWs.Cells(lastRow, 1).Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
                lastRow = lastRow + UBound(dataArray, 1)
            End If
        End If
        
        If Not wbTemp Is Nothing Then wbTemp.Close False
        
NextFile:
    Next i

    Application.ScreenUpdating = True
    
    ' Final summary
    If lastRow > 2 Then
        MsgBox "Files merged successfully into sheet '" & outputWs.Name & "'." & vbCrLf & _
               "Total rows merged: " & (lastRow - 2), vbInformation
    Else
        MsgBox "No data was merged. Please check your files and try again.", vbExclamation
    End If
End Sub

' Helper function to create a backup of the original file
Sub CreateBackup(filePath As String)
    Dim backupPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = Left(filePath, InStrRev(filePath, ".") - 1) & "_backup_" & Format(Now, "yyyymmdd_hhmmss") & Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    
    On Error Resume Next
    fso.CopyFile filePath, backupPath
    On Error GoTo 0
    
    Set fso = Nothing
End Sub