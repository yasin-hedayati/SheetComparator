Attribute VB_Name = "Module1"
Sub CompareTwoCommonSheets()
    Dim wsCommon As Worksheet
    Dim wsCommon1 As Worksheet
    Dim wsResults As Worksheet
    Dim lastRowCommon As Long
    Dim lastRowCommon1 As Long
    Dim dictCommon1 As Object
    Dim i As Long
    Dim resultRow As Long
    Dim keyVal As Variant
    Dim valCommon_B As Variant
    Dim valCommon1_B As Variant
    Dim valCommon_C As Variant
    Dim valCommon1_C As Variant
    Dim valCommon_E As Variant
    Dim valCommon1_H As Variant
    Dim valCommon_N As Variant
    Dim valCommon1_I As Variant
    Dim valCommon_I As Variant
    Dim valCommon1_E As Variant
    Dim valCommon_L As Variant
    Dim valCommon1_F As Variant
    Dim foundRow As Long
    
    ' Set sheets
    On Error Resume Next
    Set wsCommon = ThisWorkbook.Sheets("Common")
    Set wsCommon1 = ThisWorkbook.Sheets("Common1")
    On Error GoTo 0
    
    ' Check if sheets exist
    If wsCommon Is Nothing Then
        MsgBox "Sheet 'Common' not found!", vbCritical
        Exit Sub
    End If
    
    If wsCommon1 Is Nothing Then
        MsgBox "Sheet 'Common1' not found!", vbCritical
        Exit Sub
    End If
    
    ' Create results sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("ComparisonResult").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = "ComparisonResult"
    
    ' Get last rows
    lastRowCommon = wsCommon.Cells(wsCommon.Rows.Count, "A").End(xlUp).Row
    lastRowCommon1 = wsCommon1.Cells(wsCommon1.Rows.Count, "A").End(xlUp).Row
    
    ' Build dictionary from Common1 sheet (key = Column A, value = row number)
    Set dictCommon1 = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRowCommon1
        keyVal = wsCommon1.Cells(i, 1).Value
        If keyVal <> "" Then
            dictCommon1(keyVal) = i
        End If
    Next i
    
    ' Set headers in results sheet
    With wsResults
        .Cells(1, 1) = "Column A (Key)"
        .Cells(1, 2) = "Common - Col B"
        .Cells(1, 3) = "Common1 - Col B"
        .Cells(1, 4) = "Status B vs B"
        .Cells(1, 5) = "Common - Col C"
        .Cells(1, 6) = "Common1 - Col C"
        .Cells(1, 7) = "Status C vs C"
        .Cells(1, 8) = "Common - Col E"
        .Cells(1, 9) = "Common1 - Col H"
        .Cells(1, 10) = "Status E vs H"
        .Cells(1, 11) = "Common - Col N"
        .Cells(1, 12) = "Common1 - Col I"
        .Cells(1, 13) = "Status N vs I"
        .Cells(1, 14) = "Common - Col I"
        .Cells(1, 15) = "Common1 - Col E"
        .Cells(1, 16) = "Status I vs E"
        .Cells(1, 17) = "Common - Col L"
        .Cells(1, 18) = "Common1 - Col F"
        .Cells(1, 19) = "Status L vs F"
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(200, 200, 200)
    End With
    
    resultRow = 2
    
    ' Loop through Common sheet and compare with Common1
    For i = 2 To lastRowCommon
        keyVal = wsCommon.Cells(i, 1).Value
        valCommon_B = wsCommon.Cells(i, 2).Value      ' Column B
        valCommon_C = wsCommon.Cells(i, 3).Value      ' Column C
        valCommon_E = wsCommon.Cells(i, 5).Value      ' Column E
        valCommon_N = wsCommon.Cells(i, 14).Value     ' Column N
        valCommon_I = wsCommon.Cells(i, 9).Value      ' Column I
        valCommon_L = wsCommon.Cells(i, 12).Value     ' Column L
        
        ' Check if key exists in Common1
        If dictCommon1.exists(keyVal) Then
            foundRow = dictCommon1(keyVal)
            valCommon1_B = wsCommon1.Cells(foundRow, 2).Value   ' Column B
            valCommon1_C = wsCommon1.Cells(foundRow, 3).Value   ' Column C
            valCommon1_H = wsCommon1.Cells(foundRow, 8).Value   ' Column H
            valCommon1_I = wsCommon1.Cells(foundRow, 9).Value   ' Column I
            valCommon1_E = wsCommon1.Cells(foundRow, 5).Value   ' Column E
            valCommon1_F = wsCommon1.Cells(foundRow, 6).Value   ' Column F
            
            ' Write key and values to results
            wsResults.Cells(resultRow, 1) = keyVal
            wsResults.Cells(resultRow, 2) = valCommon_B
            wsResults.Cells(resultRow, 3) = valCommon1_B
            wsResults.Cells(resultRow, 5) = valCommon_C
            wsResults.Cells(resultRow, 6) = valCommon1_C
            wsResults.Cells(resultRow, 8) = valCommon_E
            wsResults.Cells(resultRow, 9) = valCommon1_H
            wsResults.Cells(resultRow, 11) = valCommon_N
            wsResults.Cells(resultRow, 12) = valCommon1_I
            wsResults.Cells(resultRow, 14) = valCommon_I
            wsResults.Cells(resultRow, 15) = valCommon1_E
            wsResults.Cells(resultRow, 17) = valCommon_L
            wsResults.Cells(resultRow, 18) = valCommon1_F
            
            ' Compare 1: Column B vs Column B
            If IsEmpty(valCommon_B) Or IsEmpty(valCommon1_B) Or valCommon_B = "" Or valCommon1_B = "" Then
                wsResults.Cells(resultRow, 4) = "MISSING"
                wsResults.Cells(resultRow, 4).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_B <> valCommon1_B Then
                wsResults.Cells(resultRow, 4) = "MISMATCH"
                wsResults.Cells(resultRow, 4).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 4) = "MATCH"
            End If
            
            ' Compare 2: Column C vs Column C
            If IsEmpty(valCommon_C) Or IsEmpty(valCommon1_C) Or valCommon_C = "" Or valCommon1_C = "" Then
                wsResults.Cells(resultRow, 7) = "MISSING"
                wsResults.Cells(resultRow, 7).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_C <> valCommon1_C Then
                wsResults.Cells(resultRow, 7) = "MISMATCH"
                wsResults.Cells(resultRow, 7).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 7) = "MATCH"
            End If
            
            ' Compare 3: Column E (Common) vs Column H (Common1)
            If IsEmpty(valCommon_E) Or IsEmpty(valCommon1_H) Or valCommon_E = "" Or valCommon1_H = "" Then
                wsResults.Cells(resultRow, 10) = "MISSING"
                wsResults.Cells(resultRow, 10).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_E <> valCommon1_H Then
                wsResults.Cells(resultRow, 10) = "MISMATCH"
                wsResults.Cells(resultRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 10) = "MATCH"
            End If
            
            ' Compare 4: Column N (Common) vs Column I (Common1)
            If IsEmpty(valCommon_N) Or IsEmpty(valCommon1_I) Or valCommon_N = "" Or valCommon1_I = "" Then
                wsResults.Cells(resultRow, 13) = "MISSING"
                wsResults.Cells(resultRow, 13).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_N <> valCommon1_I Then
                wsResults.Cells(resultRow, 13) = "MISMATCH"
                wsResults.Cells(resultRow, 13).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 13) = "MATCH"
            End If
            
            ' Compare 5: Column I (Common) vs Column E (Common1)
            If IsEmpty(valCommon_I) Or IsEmpty(valCommon1_E) Or valCommon_I = "" Or valCommon1_E = "" Then
                wsResults.Cells(resultRow, 16) = "MISSING"
                wsResults.Cells(resultRow, 16).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_I <> valCommon1_E Then
                wsResults.Cells(resultRow, 16) = "MISMATCH"
                wsResults.Cells(resultRow, 16).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 16) = "MATCH"
            End If
            
            ' Compare 6: Column L (Common) vs Column F (Common1)
            If IsEmpty(valCommon_L) Or IsEmpty(valCommon1_F) Or valCommon_L = "" Or valCommon1_F = "" Then
                wsResults.Cells(resultRow, 19) = "MISSING"
                wsResults.Cells(resultRow, 19).Interior.Color = RGB(0, 112, 192) ' Blue
            ElseIf valCommon_L <> valCommon1_F Then
                wsResults.Cells(resultRow, 19) = "MISMATCH"
                wsResults.Cells(resultRow, 19).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                wsResults.Cells(resultRow, 19) = "MATCH"
            End If
            
            resultRow = resultRow + 1
        End If
    Next i
    
    ' Auto-fit columns
    wsResults.Columns("A:S").AutoFit
    
    ' Show results
    If resultRow = 2 Then
        MsgBox "No matching keys found between Common and Common1 sheets!", vbInformation
    Else
        MsgBox "Comparison complete!" & vbNewLine & vbNewLine & _
               "? Compared " & (resultRow - 2) & " records" & vbNewLine & _
               "?? Blue = Missing value" & vbNewLine & _
               "?? Red = Value mismatch", vbInformation, "Complete"
    End If
    
End Sub

