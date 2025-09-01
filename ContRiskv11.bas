' Global variable for initial directory
Public Const INITIAL_DIRECTORY As String = "C:\Users\abc\Downloads\MMDump\"

Sub ImportMultipleTSVAndStackData()

    Dim ws As Worksheet
    Dim filePath As Variant
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long, j As Long
    Dim key As Variant
    
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
        .InitialFileName = INITIAL_DIRECTORY
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            Set ws = ActiveSheet
            ws.UsedRange.ClearContents
            
            Dim lastRow As Long
            lastRow = 1 ' Start from row 1 after clearing contents
            
            ' Loop through each selected file and stack the data
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
                
                Dim startSCNRates As Boolean, endSCNRates As Boolean
                Dim startRiskCashflow As Boolean, endRiskCashflow As Boolean
                startSCNRates = False
                endSCNRates = False
                startRiskCashflow = False
                endRiskCashflow = False
                
                ' Process data
                For i = 0 To UBound(rows)
                    If Len(Trim(rows(i))) > 0 Then
                        rowData = Application.Trim(rows(i))
                        
                        If Not foundClient And InStr(rowData, "Client:") > 0 Then
                            clientID = Trim(Split(rowData, "Client:")(1))
                            foundClient = True
                        End If
                        
                        If InStr(rowData, "Cover Ratio") > 0 Then
                            coverRatio = CDbl(Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab)))))
                            foundCover = True
                        End If
                        
                        If UCase(rowData) Like "B. SCN RATES*" Then startSCNRates = True
                        If UCase(rowData) Like "C. SCN BREAKDOWN*" Then endSCNRates = True
                        
                        If startSCNRates And Not endSCNRates And InStr(1, rowData, "FX.Rate.", vbTextCompare) > 0 And InStr(1, rowData, ".Spot", vbTextCompare) > 0 Then
                            Dim lineData As Variant
                            lineData = Split(rowData, vbTab)
                            Dim ccy As String
                            ccy = Split(lineData(0), ".")(2)
                            If IsNumeric(lineData(UBound(lineData))) Then
                                If Not fxRates.Exists(ccy) Then
                                    fxRates.Add ccy, CDbl(lineData(UBound(lineData)))
                                End If
                            End If
                        End If
                        
                        If rowData Like "K. RISK CASHFLOW*" Then startRiskCashflow = True
                        If rowData Like "L. SEPARATED DIGITAL*" Then endRiskCashflow = True
                        
                        If startRiskCashflow And Not endRiskCashflow And rowData Like "Total*" Then
                            lineData = Split(rowData, vbTab)
                            If UBound(lineData) >= 6 Then ' Ensure we have enough columns
                                Dim totalRow(2) As Variant
                                totalRow(0) = lineData(2)
                                totalRow(1) = lineData(4)
                                totalRow(2) = lineData(6)
                                dataCollection.Add totalRow
                            End If
                        End If
                    End If
                Next i
                
                ' --- 3. Write Data Vertically ---
                
                ' Add spacing between files (except for the first file)
                If lastRow > 1 Then
                    lastRow = lastRow + 2
                End If
                
                ' Write Client ID and Cover Ratio block
                ws.Cells(lastRow, 1).Value = "Client ID:"
                ws.Cells(lastRow, 2).Value = clientID
                ws.Cells(lastRow + 1, 1).Value = "Cover Ratio:"
                ws.Cells(lastRow + 1, 2).Value = coverRatio
                lastRow = lastRow + 3
                
                ' Write "Total" row data block
                If dataCollection.Count > 0 Then
                    ws.Cells(lastRow, 1).Resize(1, 3).Value = Array("CcyPair", "RiskCCy", "Exposure (RiskCCy)")
                    ws.Cells(lastRow, 1).Resize(1, 3).Font.Bold = True
                    
                    Dim totalData() As Variant
                    ReDim totalData(1 To dataCollection.Count, 1 To 3)
                    For i = 1 To dataCollection.Count
                        For j = 0 To 2
                            totalData(i, j + 1) = dataCollection(i)(j)
                        Next j
                    Next i
                    ws.Cells(lastRow + 1, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                    lastRow = lastRow + UBound(totalData, 1) + 2
                Else
                    lastRow = lastRow + 1
                End If
                
                ' Write FX rate data block
                If fxRates.Count > 0 Then
                    ws.Cells(lastRow, 1).Resize(1, 2).Value = Array("Currency", "Mid Spot Rate")
                    ws.Cells(lastRow, 1).Resize(1, 2).Font.Bold = True
                    
                    Dim fxData() As Variant
                    ReDim fxData(1 To fxRates.Count, 1 To 2)
                    i = 1
                    For Each key In fxRates.Keys
                        fxData(i, 1) = key
                        fxData(i, 2) = fxRates(key)
                        i = i + 1
                    Next key
                    ws.Cells(lastRow + 1, 1).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
                    lastRow = lastRow + UBound(fxData, 1) + 2
                Else
                    lastRow = lastRow + 1
                End If
                
            Next filePath
            
            ' Final formatting
            ws.UsedRange.Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    
    MsgBox "Data import and writing complete!", vbInformation
    
End Sub