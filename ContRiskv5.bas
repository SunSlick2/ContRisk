Sub ImportMultipleTSVAndWriteData()

    Dim ws As Worksheet
    Dim filePath As Variant
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long, j As Long
    Dim currentColumn As Long
    
    ' Data structures for the blocks of data
    Dim dataCollection As Collection
    Dim fxRates As Object
    Dim clientID As String
    Dim coverRatio As Double
    Dim foundClient As Boolean, foundCover As Boolean
    
    ' --- 1. Multiple File Selection ---
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select TSV Files"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = "C:\Users\abc\Downloads\MMDump\"
        .AllowMultiSelect = True ' Allow multiple file selection
        
        If .Show = -1 Then
            Set ws = ActiveSheet
            ws.UsedRange.ClearContents
            currentColumn = 2 ' Start writing data in column B
            
            ' Loop through each selected file
            For Each filePath In .SelectedItems
                ' Reset variables for each new file
                Set dataCollection = New Collection
                Set fxRates = CreateObject("Scripting.Dictionary")
                clientID = ""
                coverRatio = 0
                foundClient = False
                foundCover = False
                
                ' --- 2. Read and Process Data from a single file ---
                
                Open filePath For Input As #1
                fileData = Input$(LOF(1), 1)
                Close #1
                
                rows = Split(fileData, vbCrLf)
                
                For i = 0 To UBound(rows)
                    rowData = Application.Trim(rows(i))
                    
                    ' Capture Client ID
                    If Not foundClient And InStr(rowData, "Client:") > 0 Then
                        clientID = Trim(Split(rowData, "Client:")(1))
                        foundClient = True
                    End If
                    
                    ' Capture Cover Ratio
                    If InStr(rowData, "Cover Ratio") > 0 Then
                        coverRatio = CDbl(Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab)))))
                        foundCover = True
                    End If
                    
                    ' Capture FX Rates and Total Rows (same logic as before)
                    Dim startSCNRates As Boolean, endSCNRates As Boolean
                    Dim startRiskCashflow As Boolean, endRiskCashflow As Boolean
                    
                    If UCase(rowData) Like "B. SCN RATES*" Then startSCNRates = True
                    If UCase(rowData) Like "C. SCN BREAKDOWN*" Then endSCNRates = True
                    
                    If startSCNRates And Not endSCNRates And InStr(1, rowData, "FX.Rate.", vbTextCompare) > 0 And InStr(1, rowData, ".Spot", vbTextCompare) > 0 Then
                        Dim lineData As Variant
                        lineData = Split(rowData, vbTab)
                        Dim ccy As String
                        ccy = Split(lineData(0), ".")(2)
                        If IsNumeric(lineData(UBound(lineData))) Then fxRates.Add ccy, CDbl(lineData(UBound(lineData)))
                    End If
                    
                    If rowData Like "K. RISK CASHFLOW*" Then startRiskCashflow = True
                    If rowData Like "L. SEPARATED DIGITAL*" Then endRiskCashflow = True
                    
                    If startRiskCashflow And Not endRiskCashflow And rowData Like "Total*" Then
                        lineData = Split(rowData, vbTab)
                        Dim totalRow(2) As Variant
                        totalRow(0) = lineData(2) ' CcyPair
                        totalRow(1) = lineData(4) ' RiskCCy
                        totalRow(2) = lineData(6) ' Exposure (RiskCCy)
                        dataCollection.Add totalRow
                    End If
                Next i
                
                ' --- 3. Write Data Side-by-Side ---
                
                ' Write headers only once for the first file
                If currentColumn = 2 Then
                    ws.Cells(1, 1).Value = "Client ID"
                    ws.Cells(2, 1).Value = "Cover Ratio"
                    ws.Cells(4, 1).Value = "CcyPair"
                    ws.Cells(4, 2).Value = "RiskCCy"
                    ws.Cells(4, 3).Value = "Exposure (RiskCCy)"
                    ws.Cells(4 + dataCollection.Count + 2, 1).Value = "Currency"
                    ws.Cells(4 + dataCollection.Count + 2, 2).Value = "Mid Spot Rate"
                End If
                
                ' Write Client ID and Cover Ratio
                ws.Cells(1, currentColumn).Value = clientID
                ws.Cells(2, currentColumn).Value = coverRatio
                
                ' Write "Total" row data block
                If dataCollection.Count > 0 Then
                    Dim totalData() As Variant
                    ReDim totalData(1 To dataCollection.Count, 1 To 3)
                    For i = 1 To dataCollection.Count
                        For j = 0 To 2
                            totalData(i, j + 1) = dataCollection(i)(j)
                        Next j
                    Next i
                    ws.Cells(5, currentColumn).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                End If
                
                ' Write FX rate data block
                If fxRates.Count > 0 Then
                    Dim fxData() As Variant
                    ReDim fxData(1 To fxRates.Count, 1 To 2)
                    i = 1
                    For Each key In fxRates.Keys
                        fxData(i, 1) = key
                        fxData(i, 2) = fxRates(key)
                        i = i + 1
                    Next key
                    ws.Cells(4 + dataCollection.Count + 3, currentColumn).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
                End If
                
                ' Move to the next block of columns for the next file
                currentColumn = currentColumn + 4 ' Adjust this value as needed based on data width
            Next filePath
            
            ' Final formatting
            ws.UsedRange.Columns.AutoFit
            ws.Rows(1).Font.Bold = True
            ws.Rows(2).Font.Bold = True
            ws.Rows(4).Font.Bold = True
            ws.Rows(4 + dataCollection.Count + 2).Font.Bold = True
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    
    MsgBox "Data import and writing complete!", vbInformation
    
End Sub