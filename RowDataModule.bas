Attribute VB_Name = "RowDataModule"
Dim positionCurrency_ As String
Dim fundCode_ As String

Dim inputSheet As Worksheet
Dim resultSheet As Worksheet

Sub main()
    
    positionCurrency_ = Application.InputBox("Position Currency", "Select position currency", Type:=2)
    fundCode_ = Application.InputBox("Fund Code", "Select fund Code", Type:=2)
    
    
    Set inputSheet = Worksheets(1)
    CreateResultSheet "Results"
    Set resultSheet = Worksheets("Results")
    PrepareSheetHeaders resultSheet
    
    Dim resultRowNumber As Long
    resultRowNumber = 2
        
    Dim i As Long
    For i = 2 To 40000
        Dim row As rowData
        
        If inputSheet.Range("A" & i).Value <> "Excluded" Then
            If IsPositionMatching(i, positionCurrency_) And isFundCodeMatching(i, fundCode_) Then
                Set row = GetRowData(i)

                WriteToSheet resultSheet, resultRowNumber, row
                resultRowNumber = resultRowNumber + 1
            End If
        End If
    Next
End Sub

' Create rowData object from given row
Private Function GetRowData(rowNumber As Long) As rowData
    Dim row As New rowData
    With row
        .Status = inputSheet.Range("A" & rowNumber).Value
        .BasicDate = inputSheet.Range("N" & rowNumber).Value
        .FundCode = inputSheet.Range("P" & rowNumber).Value
        .BFTAccount = inputSheet.Range("Q" & rowNumber).Value
        .PositionCurrency = inputSheet.Range("R" & rowNumber).Value
        .BreakMGM = inputSheet.Range("Z" & rowNumber).Value
    End With
      
    Set GetRowData = row
End Function

' Check if position input is matching
Private Function IsPositionMatching(row As Long, Position As String) As Boolean
    IsPositionMatching = False
    If inputSheet.Range("R" & row).Value = Position Then
        IsPositionMatching = True
    End If
End Function

' Check if fund code input is matching
Private Function isFundCodeMatching(row As Long, FundCode As String) As Boolean
    isFundCodeMatching = False
    If inputSheet.Range("P" & row).Value = FundCode Then
        isFundCodeMatching = True
    End If
End Function

' Create results sheet if it does not exists already
Private Sub CreateResultSheet(sheetName As String)
    Dim sheet As Worksheet
    Dim exists As Boolean
    
    For Each sheet In Worksheets
        If sheet.Name = sheetName Then
            exists = True
        End If
    Next
    
    If Not exists Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = sheetName
    End If
End Sub

Private Sub WriteToSheet(sheet As Worksheet, rowNumber As Long, rowData As rowData)
    sheet.Select
    sheet.Range("A" & rowNumber).Value = rowData.FundCode
    sheet.Range("B" & rowNumber).Value = rowData.BFTAccount
    sheet.Range("C" & rowNumber).Value = rowData.PositionCurrency
    sheet.Range("D" & rowNumber).Value = rowData.BasicDate
    sheet.Range("E" & rowNumber).Value = rowData.BreakMGM
End Sub

Private Sub PrepareSheetHeaders(sheet As Worksheet)
    sheet.Select
    sheet.Range("A1").Value = "Fund Code"
    sheet.Range("B1").Value = "BFT Account"
    sheet.Range("C1").Value = "Position Currency"
    sheet.Range("D1").Value = "Basic Date"
    sheet.Range("E1").Value = "Break MGM"
End Sub

