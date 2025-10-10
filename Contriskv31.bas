' Global variable for initial directory
Public Const INITIAL_DIRECTORY As String = "C:\Users\abc\Downloads\MMDump\"

Sub ImportMultipleTSVAndStackData()
    Dim ws As Worksheet
    Dim filePath As Variant
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long, k As Long
    Dim key As Variant
    
    ' Variable declarations
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
    
    ' Variables for the new in-place update logic
    Dim fileMap As Object
    Dim currentRow As Long
    Dim clientIDValue As String
    Dim rowsWritten As Long
    Dim nextClientIDRow As Long
    Dim oldBlockEndRow As Long

    ' Data structures for the blocks of data
    Dim dataCollection As Collection
    Dim fxRates As Object
    Dim clientID As String ' ID extracted from filename
    Dim coverRatio As Double
    Dim foundCover As Boolean

    ' --- 1. Setup and File Selection ---
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set fileMap = CreateObject("Scripting.Dictionary") ' Used to map ClientID to file path

    With fd
        .Title = "Select TSV Files"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = INITIAL_DIRECTORY
        .AllowMultiSelect = True

        If .Show = -1 Then
            ' Set the output worksheet to "ContRisk2"
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets("ContRisk2")
            On Error GoTo 0
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.Name = "ContRisk2"
            End If
            
            ' Headers are assumed to be in Row 1. We will not clear the sheet.
            
            ' Pre-load the ToFlipTable from columns M, N, O
            Set ToFlipTable = New Collection
            k = 2 ' Start from row 2 to skip the header
            Dim wsCurrent As Worksheet ' Use a fresh worksheet object for reading M:O
            Set wsCurrent = ThisWorkbook.Sheets(ws.Name) 
            
            Do While Len(wsCurrent.Cells(k, 13).Value) > 0 And (Len(wsCurrent.Cells(k, 14).Value) > 0 Or Len(wsCurrent.Cells(k, 15).Value) > 0)
                flipEntry(0) = wsCurrent.Cells(k, 13).Value ' Currency pair from column M
                flipEntry(1) = wsCurrent.Cells(k, 14).Value ' Divisor currency from column N
                flipEntry(2) = wsCurrent.Cells(k, 15).Value ' Multiplier currency from column O
                ToFlipTable.Add flipEntry
                k = k + 1
            Loop
            
            ' --- 2. Map Selected TSV Files to Client IDs ---
            For Each filePath In .SelectedItems
                ' Extract ClientID from Filename
                Dim fileName As String
                Dim startPos As Long
                Dim endPos As Long
                
                fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                startPos = InStr(fileName, "PORTFOLIO_ANALYSIS_") + Len("PORTFOLIO_ANALYSIS_")
                endPos = InStr(startPos, fileName, "_")
                
                If startPos > Len("PORTFOLIO_ANALYSIS_") And endPos > startPos Then
                    clientID = Mid(fileName, startPos, endPos - startPos)
                Else
                    clientID = "N/A"
                End If
                
                If Not fileMap.Exists(clientID) Then
                    fileMap.Add clientID, filePath
                End If
            Next filePath
            
            ' --- 3. Main Loop: Iterate Column A and Update Data In-Place ---

            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            currentRow = 2 ' Start after the header
            
            Do While currentRow <= lastRow
                clientIDValue = Trim(ws.Cells(currentRow, 1).Value)

                ' Only update if a valid ClientID is found in Col A AND we have a file for it
                If Len(clientIDValue) > 0 And fileMap.Exists(clientIDValue) Then
                    filePath = fileMap(clientIDValue)
                    
                    ' Reset variables for processing the TSV file
                    Set dataCollection = New Collection
                    Set fxRates = CreateObject("Scripting.Dictionary")
                    coverRatio = 0
                    foundCover = False
                    
                    ' --- Read and Process Data from TSV file ---
                    Open filePath For Input As #1
                    fileData = Input$(LOF(1), 1)
                    Close #1
                    
                    rows = Split(fileData, vbCrLf)
                    startSCNRates = False: endSCNRates = False
                    startRiskCashflow = False: endRiskCashflow = False
                    
                    For i = 0 To UBound(rows)
                        If Len(Trim(rows(i))) > 0 Then
                            rowData = Application.Trim(rows(i))
                            
                            ' Find Cover Ratio
                            If InStr(rowData, "Cover Ratio") > 0 Then
                                coverRatioString = Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab))))
                                coverRatioString = Replace(coverRatioString, ",", "")
                                If IsNumeric(coverRatioString) Then coverRatio = CDbl(coverRatioString)
                                foundCover = True
                            End If
                            
                            ' Find FX Rates
                            If UCase(rowData) Like "B. SCN RATES*" Then startSCNRates = True
                            If UCase(rowData) Like "C. SCN BREAKDOWN*" Then endSCNRates = True
                            If startSCNRates And Not endSCNRates And InStr(1, rowData, "FX.Rate.", vbTextCompare) > 0 And InStr(1, rowData, ".Spot", vbTextCompare) > 0 Then
                                lineData = Split(rowData, vbTab)
                                ccy = Split(lineData(0), ".")(2)
                                fxRateString = Trim(lineData(UBound(lineData)))
                                fxRateString = Replace(fxRateString, ",", "")
                                If IsNumeric(fxRateString) Then
                                    If Not fxRates.Exists(ccy) Then fxRates.Add ccy, CDbl(fxRateString)
                                End If
                            End If
                            
                            ' Find Risk Cashflow Data
                            If rowData Like "K. RISK CASHFLOW*" Then startRiskCashflow = True
                            If rowData Like "L. SEPARATED DIGITAL*" Then endRiskCashflow = True
                            If startRiskCashflow And Not endRiskCashflow And rowData Like "Total*" Then
                                lineData = Split(rowData, vbTab)
                                If UBound(lineData) >= 6 Then
                                    Dim totalRow(2) As Variant
                                    totalRow(0) = lineData(2)
                                    totalRow(1) = lineData(4)
                                    exposureString = Trim(lineData(6))
                                    exposureString = Replace(exposureString, ",", "")
                                    If IsNumeric(exposureString) Then totalRow(2) = CDbl(exposureString) Else totalRow(2) = 0
                                    dataCollection.Add totalRow
                                End If
                            End If
                        End If
                    Next i
                    
                    ' --- Prepare and Write Data In-Place ---

                    ' Determine size of the array for writing
                    If dataCollection.Count = 0 Then
                        ReDim totalData(1 To 1, 1 To 8)
                    Else
                        ReDim totalData(1 To dataCollection.Count, 1 To 8)
                    End If

                    For i = 1 To UBound(totalData, 1)
                        ' Populate ClientID and CoverRatio only for the first row
                        If i = 1 Then
                            totalData(i, 1) = clientIDValue ' Use the ID from Column A
                            totalData(i, 2) = Int(coverRatio)
                        Else
                            totalData(i, 1) = ""
                            totalData(i, 2) = ""
                        End If
                        
                        ' Populate other columns only if cashflow data exists
                        If dataCollection.Count > 0 Then
                            totalData(i, 3) = dataCollection(i)(0) ' CcyPair
                            totalData(i, 6) = dataCollection(i)(1) ' Risk Ccy
                            totalData(i, 7) = dataCollection(i)(2) ' Exposure(RiskCcy)
                            
                            ' Apply manipulation logic
                            ccyPair = dataCollection(i)(0) ' Use the original CcyPair for calculation
                            
                            For k = 1 To ToFlipTable.Count
                                currentFlipEntry = ToFlipTable(k)
                                
                                If ccyPair = currentFlipEntry(0) Then
                                    divCurrency = currentFlipEntry(1)
                                    mulCurrency = currentFlipEntry(2)
                                    originalExposure = totalData(i, 7)
                                    divRate = 1: mulRate = 1
                                    
                                    For Each key In fxRates.Keys
                                        If key = divCurrency Then divRate = fxRates(key)
                                        If key = mulCurrency Then mulRate = fxRates(key)
                                    Next key
                                    
                                    If divRate <> 0 And mulRate <> 0 Then
                                        manipulatedExposure = (originalExposure / divRate) * mulRate * -1
                                        totalData(i, 8) = manipulatedExposure
                                    End If
                                    
                                    ' DISPLAY FLIP - Flip CcyPair for display ONLY (Requirement #2)
                                    If Len(ccyPair) = 6 Then
                                        totalData(i, 3) = Right(ccyPair, 3) & Left(ccyPair, 3)
                                    End If
                                    
                                    Exit For
                                End If
                            Next k
                            
                            ' Populate the "Exposure (mio)" column
                            If Not IsEmpty(totalData(i, 8)) And IsNumeric(totalData(i, 8)) Then
                                valueForMio = CDbl(totalData(i, 8))
                            ElseIf IsNumeric(totalData(i, 7)) Then
                                valueForMio = CDbl(totalData(i, 7))
                            Else
                                valueForMio = 0
                            End If
                            totalData(i, 4) = Round(valueForMio / 1000000, 1)
                        End If
                    Next i
                    
                    ' Write main data block
                    ws.Cells(currentRow, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                    rowsWritten = UBound(totalData, 1)
                    
                    ' Write FX rate data block to the right
                    If fxRates.Count > 0 Then
                        ReDim fxData(1 To fxRates.Count, 1 To 2)
                        i = 1
                        For Each key In fxRates.Keys
                            fxData(i, 1) = key
                            fxData(i, 2) = fxRates(key)
                            i = i + 1
                        Next key
                        ' Start writing FX data at the same row as the main block data
                        ws.Cells(currentRow, 10).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
                        rowsWritten = Application.Max(rowsWritten, UBound(fxData, 1))
                    End If
                    
                    ' --- Clean up old data and set cursor for next loop ---
                    
                    ' Find the row where the next ClientID is listed (or end of sheet)
                    nextClientIDRow = currentRow + 1
                    Do While nextClientIDRow <= lastRow And Len(Trim(ws.Cells(nextClientIDRow, 1).Value)) = 0
                        nextClientIDRow = nextClientIDRow + 1
                    Loop
                    
                    ' Calculate the area of old data to clear
                    oldBlockEndRow = nextClientIDRow - 1
                    
                    ' Clear old content in columns A:L if the new data is shorter
                    If currentRow + rowsWritten <= oldBlockEndRow Then
                        ws.Range(ws.Cells(currentRow + rowsWritten, 1), ws.Cells(oldBlockEndRow, 12)).ClearContents
                    End If
                    
                    ' Set currentRow to the row of the next ClientID found
                    currentRow = nextClientIDRow
                
                Else
                    ' ClientID is not empty, but we don't have a matching file (or it's an empty cell)
                    currentRow = currentRow + 1
                End If
            Loop
            
            ' Final formatting
            ws.Range("A:L").Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    
    MsgBox "Data update complete on sheet 'ContRisk2'!", vbInformation
    
End Sub
