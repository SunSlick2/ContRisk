' Global variable for initial directory
Public Const INITIAL_DIRECTORY As String = "C:\Users\abc\Downloads\MMDump\"

Sub ImportMultipleTSVAndStackData()
    Dim ws As Worksheet
    Dim filePath As Variant
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long, j As Long, k As Long
    Dim key As Variant
    
    ' Variable declarations moved to the top of the procedure to avoid duplicate declaration errors
    Dim fd As FileDialog
    Dim lastRow As Long
    Dim ToFlipTable As Collection
    Dim flipEntry(2) As String
    Dim startSCNRates As Boolean, endSCNRates As Boolean
    Dim startRiskCashflow As Boolean, endRiskCashflow As Boolean
    Dim lineData As Variant
    Dim ccy As String
    Dim totalData() As Variant
    Dim ccyPair As String
    Dim currentFlipEntry() As String
    Dim divCurrency As String
    Dim mulCurrency As String
    Dim originalExposure As Double
    Dim manipulatedExposure As Double
    Dim divRate As Double
    Dim mulRate As Double
    Dim valueForMio As Double
    Dim fxData() As Variant
    Dim coverRatioString As String ' Variable for handling Cover Ratio string
    Dim exposureString As String ' Variable for handling Exposure string
    Dim fxRateString As String ' Variable for handling FX Rate string
    Dim rowsInMainTable As Long
    Dim rowsInFxTable As Long

    ' Data structures for the blocks of data
    Dim dataCollection As Collection
    Dim fxRates As Object
    Dim clientID As String
    Dim coverRatio As Double
    Dim foundClient As Boolean, foundCover As Boolean

    ' --- 1. Multiple File Selection ---
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select TSV Files"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = INITIAL_DIRECTORY
        .AllowMultiSelect = True

        If .Show = -1 Then
            Set ws = ActiveSheet

            ' Clear only columns A through L (preserve M, N, O manipulation parameters)
            ws.Range("A:L").Clear
            
            lastRow = 1 ' Start from row 1 after clearing contents

            ' Pre-load the ToFlipTable from columns M, N, O
            Set ToFlipTable = New Collection
            k = 2 ' Start from row 2 to skip the header

            ' The corrected loop condition to ensure all necessary columns have data
            Do While Len(ws.Cells(k, 13).Value) > 0 And (Len(ws.Cells(k, 14).Value) > 0 Or Len(ws.Cells(k, 15).Value) > 0)
                flipEntry(0) = ws.Cells(k, 13).Value ' Currency pair from column M
                flipEntry(1) = ws.Cells(k, 14).Value ' Divisor currency from column N
                flipEntry(2) = ws.Cells(k, 15).Value ' Multiplier currency from column O
                ToFlipTable.Add flipEntry
                k = k + 1
            Loop
            
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
                            coverRatioString = Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab))))
                            coverRatioString = Replace(coverRatioString, ",", "") ' Clean up value
                            If IsNumeric(coverRatioString) Then
                                coverRatio = CDbl(coverRatioString)
                            End If
                            foundCover = True
                        End If
                        
                        If UCase(rowData) Like "B. SCN RATES*" Then startSCNRates = True
                        If UCase(rowData) Like "C. SCN BREAKDOWN*" Then endSCNRates = True
                        
                        If startSCNRates And Not endSCNRates And InStr(1, rowData, "FX.Rate.", vbTextCompare) > 0 And InStr(1, rowData, ".Spot", vbTextCompare) > 0 Then
                            lineData = Split(rowData, vbTab)
                            ccy = Split(lineData(0), ".")(2)
                            fxRateString = Trim(lineData(UBound(lineData)))
                            fxRateString = Replace(fxRateString, ",", "")
                            If IsNumeric(fxRateString) Then
                                If Not fxRates.Exists(ccy) Then
                                    fxRates.Add ccy, CDbl(fxRateString)
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
                                exposureString = Trim(lineData(6))
                                exposureString = Replace(exposureString, ",", "")
                                If IsNumeric(exposureString) Then
                                    totalRow(2) = CDbl(exposureString)
                                Else
                                    totalRow(2) = 0
                                End If
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
                
                Dim startRowOfMainTable As Long
                startRowOfMainTable = lastRow
                
                ' Write "Total" row data block
                If dataCollection.Count > 0 Then
                    ws.Cells(lastRow, 1).Resize(1, 7).Value = Array("ClientID", "CoverRatio", "CcyPair", "Exposure(mio)", "Risk Ccy", "Exposure(RiskCcy)", "Manipulated Exposure")
                    
                    ReDim totalData(1 To dataCollection.Count, 1 To 7)
                    For i = 1 To dataCollection.Count
                        ' Populate new columns first, only for the first row
                        If i = 1 Then
                            totalData(i, 1) = clientID
                            totalData(i, 2) = Int(coverRatio)
                        Else
                            totalData(i, 1) = ""
                            totalData(i, 2) = ""
                        End If
                        
                        ' Populate other columns based on the new layout
                        totalData(i, 3) = dataCollection(i)(0) ' CcyPair
                        totalData(i, 5) = dataCollection(i)(1) ' Risk Ccy
                        totalData(i, 6) = dataCollection(i)(2) ' Exposure(RiskCcy)
                        
                    ' Apply manipulation logic if ccypair appears in ToFlipTable
                    ccyPair = totalData(i, 3) ' Use the value from the new array
                    
                    ' Check if this currency pair exists in the ToFlipTable
                    For k = 1 To ToFlipTable.Count
                        currentFlipEntry = ToFlipTable(k) ' This works for retrieving arrays from collections
                        
                        If ccyPair = currentFlipEntry(0) Then
                            divCurrency = currentFlipEntry(1) ' Divisor currency from ToFlipTable
                            mulCurrency = currentFlipEntry(2) ' Multiplier currency from ToFlipTable
                            
                            originalExposure = totalData(i, 6) ' Already a number
                            
                            ' Find the FX rates for the specified currencies
                            divRate = 1
                            mulRate = 1
                            
                            For Each key In fxRates.Keys
                                If key = divCurrency Then divRate = fxRates(key)
                                If key = mulCurrency Then mulRate = fxRates(key)
                            Next key
                            
                            If divRate <> 0 And mulRate <> 0 Then
                                manipulatedExposure = (originalExposure / divRate) * mulRate * -1
                                totalData(i, 7) = manipulatedExposure
                            End If
                            
                            Exit For ' Exit loop once we found a match
                        End If
                    Next k
                    
                    ' Populate the "Exposure (mio)" column
                    If Not IsEmpty(totalData(i, 7)) And IsNumeric(totalData(i, 7)) Then
                        valueForMio = CDbl(totalData(i, 7))
                    ElseIf IsNumeric(totalData(i, 6)) Then
                        valueForMio = CDbl(totalData(i, 6))
                    Else
                        valueForMio = 0
                    End If

                    totalData(i, 4) = Round(valueForMio / 1000000, 1)

                Next i
                    
                    ws.Cells(lastRow + 1, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                    rowsInMainTable = UBound(totalData, 1) + 1 ' +1 for the header row
                Else
                    rowsInMainTable = 0
                End If
                
                ' Write FX rate data block to the right
                If fxRates.Count > 0 Then
                    ws.Cells(startRowOfMainTable, 9).Resize(1, 2).Value = Array("Currency", "Mid Spot Rate")
                    
                    ReDim fxData(1 To fxRates.Count, 1 To 2)
                    i = 1
                    For Each key In fxRates.Keys
                        fxData(i, 1) = key
                        fxData(i, 2) = fxRates(key)
                        i = i + 1
                    Next key
                    ws.Cells(startRowOfMainTable + 1, 9).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
                    rowsInFxTable = UBound(fxData, 1) + 1 ' +1 for the header row
                Else
                    rowsInFxTable = 0
                End If
                
                ' Update lastRow for the next file
                lastRow = lastRow + Application.Max(rowsInMainTable, rowsInFxTable) + 2
                
            Next filePath
            
            ' Final formatting - only autofit columns A through L
            ws.Range("A:L").Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    
    MsgBox "Data import and writing complete!", vbInformation
    
End Sub
