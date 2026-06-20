Attribute VB_Name = "Module2"
Sub ExtractMismatchRecords()
    Dim wsResults As Worksheet
    Dim wsMismatch As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim mismatchRow As Long
    Dim hasMismatch As Boolean
    Dim mismatchColumns As Variant
    
    ' Set the comparison results sheet
    On Error Resume Next
    Set wsResults = ThisWorkbook.Sheets("ComparisonResult")
    On Error GoTo 0
    
    ' Check if ComparisonResult sheet exists
    If wsResults Is Nothing Then
        MsgBox "Sheet 'ComparisonResult' not found! Please run the comparison first.", vbCritical
        Exit Sub
    End If
    
    ' Columns to check for mismatch (D, G, J, M, P, S)
    mismatchColumns = Array(4, 7, 10, 13, 16, 19) ' Column numbers
    
    ' Get last row and column
    lastRow = wsResults.Cells(wsResults.Rows.Count, "A").End(xlUp).Row
    lastCol = wsResults.Cells(1, wsResults.Columns.Count).End(xlToLeft).Column
    
    ' Check if there's data beyond headers
    If lastRow < 2 Then
        MsgBox "No data found in ComparisonResult sheet!", vbInformation
        Exit Sub
    End If
    
    ' Create mismatch sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("MismatchRecords").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsMismatch = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsMismatch.Name = "MismatchRecords"
    
    ' Copy headers from ComparisonResult to Mismatch sheet
    For j = 1 To lastCol
        wsMismatch.Cells(1, j).Value = wsResults.Cells(1, j).Value
    Next j
    
    ' Add an extra column to show which columns had mismatches
    wsMismatch.Cells(1, lastCol + 1).Value = "Mismatch Columns"
    wsMismatch.Cells(1, lastCol + 1).Font.Bold = True
    wsMismatch.Cells(1, lastCol + 1).Interior.Color = RGB(255, 255, 0) ' Yellow
    
    mismatchRow = 2
    
    ' Loop through each row in ComparisonResult
    For i = 2 To lastRow
        hasMismatch = False
        Dim mismatchList As String
        mismatchList = ""
        
        ' Check each mismatch column
        For j = 0 To UBound(mismatchColumns)
            Dim colNum As Long
            colNum = mismatchColumns(j)
            
            ' Check if cell contains "MISMATCH"
            If wsResults.Cells(i, colNum).Value = "MISMATCH" Then
                hasMismatch = True
                ' Add column letter to mismatch list
                Dim colLetter As String
                colLetter = Split(wsResults.Cells(1, colNum).Address, "$")(1)
                If mismatchList = "" Then
                    mismatchList = colLetter
                Else
                    mismatchList = mismatchList & ", " & colLetter
                End If
            End If
        Next j
        
        ' If mismatch found, copy entire row to mismatch sheet
        If hasMismatch Then
            ' Copy all columns from ComparisonResult
            For j = 1 To lastCol
                wsMismatch.Cells(mismatchRow, j).Value = wsResults.Cells(i, j).Value
            Next j
            
            ' Add the mismatch columns list
            wsMismatch.Cells(mismatchRow, lastCol + 1).Value = mismatchList
            
            ' Optional: Highlight the entire row in light red
            wsMismatch.Rows(mismatchRow).Interior.Color = RGB(255, 200, 200)
            
            mismatchRow = mismatchRow + 1
        End If
    Next i
    
    ' Auto-fit columns
    wsMismatch.Columns.AutoFit
    
    ' Show results
    If mismatchRow = 2 Then
        MsgBox "No mismatch records found!", vbInformation
    Else
        MsgBox "Mismatch extraction complete!" & vbNewLine & vbNewLine & _
               "? Found " & (mismatchRow - 2) & " records with mismatches" & vbNewLine & _
               "?? Check 'MismatchRecords' sheet for details", vbInformation, "Complete"
    End If
    
End Sub

