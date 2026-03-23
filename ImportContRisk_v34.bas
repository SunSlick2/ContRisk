Attribute VB_Name = "ImportContRisk"
' Global variable for initial directory
Public Const INITIAL_DIRECTORY As String = "C:\Users\1357963\Downloads\MMDump\"

Sub ImportMultipleTSVAndStackDatav3()
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

            lastRow = ws.Cells(ws.rows.Count, "K").End(xlUp).row
            
            Range("B4:K" & lastRow).ClearContents
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
                                    divRate = 1: mulRate = 1
                                    divCurrency = currentFlipEntry(1)
                                    mulCurrency = currentFlipEntry(2)
                                    originalExposure = totalData(i, 7)
                                    
                                    
                                    For Each key In fxRates.Keys
                                        If key = divCurrency Then divRate = fxRates(key)
                                        If key = mulCurrency Then mulRate = fxRates(key)
                                        If key = 1 Then
                                            mulRate = -1
                                            divRate = 1
                                        End If
                                        If key = -1 Then
                                            mulRate = 1
                                            divRate = 1
                                        End If
                                        
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
            lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
            Range("A" & currentRow + 2 & ":B" & lastRow).ClearContents
            Range("A" & currentRow + 3).Value = "Spot ref: Spot/Forward/Cash"
            Range("A" & currentRow + 4).Value = "USDJPY:"
            Range("A" & currentRow + 5).Value = "GBPJPY:"
            Range("A" & currentRow + 7).Value = "Spot ref: Structures"
            Range("A" & currentRow + 8).Value = "USDJPY:"
            Range("A" & currentRow + 9).Value = "GBPJPY:"
            
            With Range("B" & currentRow + 4)
                .Formula = "=K5"
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            With Range("B" & currentRow + 5)
                .Formula = "=K5*K4"
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            
            ' Final formatting
            ws.Range("A:L").Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    Range("D24").Value2 = "NOP" & Chr(10) & "($ mio)"
    MsgBox "Data update complete on sheet 'ContRisk2'!", vbInformation
    
End Sub

Sub ImportMultipleTSVAndStackDatav6()
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
    Dim startSCNSynthetic As Boolean, endSCNSynthetic As Boolean
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
    Dim coverRatioString As String
    Dim exposureString As String
    Dim fxRateString As String
    
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
    Dim scnSyntheticData As Collection ' New: Store Section D data
    Dim clientID As String
    Dim coverRatio As Double
    Dim foundCover As Boolean
    Dim singleCurrencySuppressed As Boolean ' New: Flag for suppressed netting
    Dim errorMessages As Collection ' New: Store error messages
    
    ' --- 1. Setup and File Selection ---
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set fileMap = CreateObject("Scripting.Dictionary")

    With fd
        .Title = "Select TSV Files"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = INITIAL_DIRECTORY
        .AllowMultiSelect = True

        If .Show = -1 Then
            ' Set the output worksheet to "ContRisk3"
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets("ContRisk3")
            On Error GoTo 0
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.Name = "ContRisk3"
                
                ' Set up headers for ContRisk3
                ws.Cells(1, 1).Value = "ClientID"
                ws.Cells(1, 2).Value = "CoverRatio"
                ws.Cells(1, 3).Value = "CcyPair"
                ws.Cells(1, 4).Value = "Delta (mio)"
                ws.Cells(1, 5).Value = "Synth. Exp."
                ws.Cells(1, 6).Value = "SOV"
                ws.Cells(1, 7).Value = "Risk Ccy"
                ws.Cells(1, 8).Value = "Exposure(RiskCcy)"
                ws.Cells(1, 9).Value = "Manipulated Exposure"
                ' Column J is blank
                ws.Cells(1, 11).Value = "Currency"
                ws.Cells(1, 12).Value = "Mid Spot rate"
                ws.Cells(1, 18).Value = "Errors" ' Column R
            End If
            
            ' Pre-load the ToFlipTable from columns N, O, P (new columns)
            Set ToFlipTable = New Collection
            k = 2 ' Start from row 2 to skip the header
            Dim wsCurrent As Worksheet
            Set wsCurrent = ThisWorkbook.Sheets(ws.Name)
            
            Do While Len(wsCurrent.Cells(k, 14).Value) > 0 And (Len(wsCurrent.Cells(k, 15).Value) > 0 Or Len(wsCurrent.Cells(k, 16).Value) > 0)
                flipEntry(0) = wsCurrent.Cells(k, 14).Value ' Currency pair from column N
                flipEntry(1) = wsCurrent.Cells(k, 15).Value ' Divisor currency from column O
                flipEntry(2) = wsCurrent.Cells(k, 16).Value ' Multiplier currency from column P
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

            lastRow = ws.Cells(ws.rows.Count, "K").End(xlUp).row
            
            ' Clear columns B through K (main data area)
            If lastRow > 1 Then
                ws.Range("B4:K" & lastRow).ClearContents
            End If
            
            currentRow = 2 ' Start after the header
            
            Do While currentRow <= lastRow
                clientIDValue = Trim(ws.Cells(currentRow, 1).Value)

                ' Only update if a valid ClientID is found in Col A AND we have a file for it
                If Len(clientIDValue) > 0 And fileMap.Exists(clientIDValue) Then
                    filePath = fileMap(clientIDValue)
                    
                    ' Reset variables for processing the TSV file
                    Set dataCollection = New Collection
                    Set fxRates = CreateObject("Scripting.Dictionary")
                    Set scnSyntheticData = New Collection ' New: For Section D data
                    Set errorMessages = New Collection ' New: For error logging
                    coverRatio = 0
                    foundCover = False
                    singleCurrencySuppressed = False ' New: Reset flag
                    
                    ' --- Read and Process Data from TSV file ---
                    Open filePath For Input As #1
                    fileData = Input$(LOF(1), 1)
                    Close #1
                    
                    rows = Split(fileData, vbCrLf)
                    startSCNRates = False: endSCNRates = False
                    startRiskCashflow = False: endRiskCashflow = False
                    startSCNSynthetic = False: endSCNSynthetic = False ' New: For Section D
                    
                    For i = 0 To UBound(rows)
                        If Len(Trim(rows(i))) > 0 Then
                            rowData = Application.Trim(rows(i))
                            
                            ' Check for Single Currency Netting Suppressed
                            If InStr(rowData, "<*** SINGLE CURRENCY NETTING SUPPRESSED ***>") > 0 Then
                                singleCurrencySuppressed = True
                            End If
                            
                            ' Find Cover Ratio
                            If InStr(rowData, "Cover Ratio") > 0 Then
                                coverRatioString = Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab))))
                                coverRatioString = Replace(coverRatioString, ",", "")
                                If IsNumeric(coverRatioString) Then coverRatio = CDbl(coverRatioString)
                                foundCover = True
                            End If
                            
                            ' Find FX Rates (Section B)
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
                            
                            ' Find SCN Synthetic Positions Data (Section D) - NEW
                            If rowData Like "D. SCN SYNTHETIC POSITIONS*" Then startSCNSynthetic = True
                            If rowData Like "E. SYNTHETIC POSITIONS*" Then endSCNSynthetic = True
                            If startSCNSynthetic And Not endSCNSynthetic Then
                                ' Skip divider lines and empty lines
                                rowData = Trim(rowData)
                                If Len(rowData) > 0 And Not (rowData Like "----------*" Or rowData Like "---*" Or rowData Like "----------------------------------------------------------------*") Then
                                    ' Parse the SCN Synthetic Positions data
                                    lineData = Split(rowData, vbTab)
                                    If UBound(lineData) >= 6 Then ' Need at least 7 columns for full data
                                        Dim scnRow(7) As String ' 8 columns (0-7) to hold all data
                                        For k = 0 To 7
                                            If k <= UBound(lineData) Then
                                                scnRow(k) = Trim(lineData(k))
                                            Else
                                                scnRow(k) = ""
                                            End If
                                        Next k
                                        
                                        ' Only add if first column has a currency pair (6 letters)
                                        If Len(scnRow(0)) >= 6 Then
                                            scnSyntheticData.Add scnRow
                                        End If
                                    End If
                                End If
                            End If
                            
                            ' Find Risk Cashflow Data (Section K)
                            If rowData Like "K. RISK CASHFLOW*" Then startRiskCashflow = True
                            If rowData Like "L. SEPARATED DIGITAL*" Then endRiskCashflow = True
                            If startRiskCashflow And Not endRiskCashflow And rowData Like "Total*" Then
                                lineData = Split(rowData, vbTab)
                                If UBound(lineData) >= 6 Then
                                    Dim totalRow(2) As Variant
                                    totalRow(0) = lineData(2) ' CcyPair
                                    totalRow(1) = lineData(4) ' Risk Ccy
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
                        ReDim totalData(1 To 1, 1 To 9) ' Now 9 columns (A-I)
                        
                        ' Just write basic client info if no data
                        totalData(1, 1) = clientIDValue
                        totalData(1, 2) = Int(coverRatio)
                        
                        ' Skip all the complex processing since there's no data
                        GoTo WriteData
                    Else
                        ReDim totalData(1 To dataCollection.Count, 1 To 9)
                    End If

                    ' Store display versions of CcyPairs for matching (only if we have data)
                    Dim displayCcyPairs() As String
                    ReDim displayCcyPairs(1 To dataCollection.Count)

                    For i = 1 To UBound(totalData, 1)
                        ' Populate ClientID and CoverRatio only for the first row
                        If i = 1 Then
                            totalData(i, 1) = clientIDValue ' Use the ID from Column A
                            totalData(i, 2) = Int(coverRatio)
                        Else
                            totalData(i, 1) = ""
                            totalData(i, 2) = ""
                        End If
                        
                        ' Populate other columns (we know dataCollection.Count > 0 here)
                        Dim originalCcyPair As String
                        originalCcyPair = dataCollection(i)(0) ' Original CcyPair from Section K
                        
                        ' Apply DISPLAY FLIP logic to get what's shown in column C
                        Dim displayCcyPair As String
                        displayCcyPair = originalCcyPair ' Start with original
                        
                        ' Check if this pair needs to be flipped for display
                        For k = 1 To ToFlipTable.Count
                            currentFlipEntry = ToFlipTable(k)
                            
                            If originalCcyPair = currentFlipEntry(0) Then
                                ' This pair is in the ToFlipTable, so flip it for display
                                If Len(originalCcyPair) = 6 Then
                                    displayCcyPair = Right(originalCcyPair, 3) & Left(originalCcyPair, 3)
                                End If
                                Exit For
                            End If
                        Next k
                        
                        ' Store the display version
                        displayCcyPairs(i) = displayCcyPair
                        totalData(i, 3) = displayCcyPair ' Column C shows flipped version
                        
                        totalData(i, 7) = dataCollection(i)(1) ' Risk Ccy (now column G)
                        totalData(i, 8) = dataCollection(i)(2) ' Exposure(RiskCcy) (now column H)
                        
                        ' Apply manipulation logic for exposure calculation
                        ccyPair = originalCcyPair ' Use original for calculation
                        
                        For k = 1 To ToFlipTable.Count
                            currentFlipEntry = ToFlipTable(k)
                            
                            If ccyPair = currentFlipEntry(0) Then
                                divRate = 1: mulRate = 1
                                divCurrency = currentFlipEntry(1)
                                mulCurrency = currentFlipEntry(2)
                                originalExposure = totalData(i, 8)
                                
                                For Each key In fxRates.Keys
                                    If key = divCurrency Then divRate = fxRates(key)
                                    If key = mulCurrency Then mulRate = fxRates(key)
                                    If key = 1 Then
                                        mulRate = -1
                                        divRate = 1
                                    End If
                                    If key = -1 Then
                                        mulRate = 1
                                        divRate = 1
                                    End If
                                Next key
                                
                                If divRate <> 0 And mulRate <> 0 Then
                                    manipulatedExposure = (originalExposure / divRate) * mulRate * -1
                                    totalData(i, 9) = manipulatedExposure ' Manipulated Exposure (column I)
                                End If
                                
                                Exit For
                            End If
                        Next k
                        
                        ' Populate the "Delta (mio)" column (column D)
                        If Not IsEmpty(totalData(i, 9)) And IsNumeric(totalData(i, 9)) Then
                            valueForMio = CDbl(totalData(i, 9))
                        ElseIf IsNumeric(totalData(i, 8)) Then
                            valueForMio = CDbl(totalData(i, 8))
                        Else
                            valueForMio = 0
                        End If
                        totalData(i, 4) = Round(valueForMio / 1000000, 1)
                        
                        ' Initialize Synth. Exp. column (column E) - will be populated later
                        totalData(i, 5) = 0
                    Next i
                    
                    ' --- NEW: Process SCN Synthetic Positions Data for Synthetic Exposure ---
                    If Not singleCurrencySuppressed And scnSyntheticData.Count > 0 Then
                        ' Process each currency pair from the main data
                        For i = 1 To UBound(totalData, 1)
                            ' Get the DISPLAYED CcyPair (what's shown in column C)
                            Dim displayedCcyPair As String
                            displayedCcyPair = displayCcyPairs(i)
                            
                            ' Debug: Add to error column for verification
                            Dim debugInfo As String
                            debugInfo = "Display: " & displayedCcyPair & ", Original: " & dataCollection(i)(0)
                            'debugInfo = dataCollection(i)(0)
                            ' Look for matching data in SCN Synthetic Positions
                            Dim foundMatch As Boolean
                            foundMatch = False
                            Dim matchedTSVPair As String
                            Dim wasFlippedForMatch As Boolean
                            
                            For k = 1 To scnSyntheticData.Count
                                Dim scnRowData() As String
                                scnRowData = scnSyntheticData(k)
                                
                                ' Check if we have valid data
                                If UBound(scnRowData) >= 6 Then
                                    Dim tsvCurrencyPair As String
                                    tsvCurrencyPair = scnRowData(0)
                                    
                                    ' Try to match against DISPLAYED pair
                                    ' Also try flipped version if needed
                                    Dim tsvFlippedPair As String
                                    If Len(tsvCurrencyPair) = 6 Then
                                        tsvFlippedPair = Right(tsvCurrencyPair, 3) & Left(tsvCurrencyPair, 3)
                                    Else
                                        tsvFlippedPair = ""
                                    End If
                                    
                                    ' Check for match
                                    If tsvCurrencyPair = displayedCcyPair Then
                                        foundMatch = True
                                        matchedTSVPair = tsvCurrencyPair
                                        wasFlippedForMatch = False
                                    ElseIf tsvFlippedPair = displayedCcyPair Then
                                        foundMatch = True
                                        matchedTSVPair = tsvCurrencyPair ' Keep original TSV pair
                                        wasFlippedForMatch = True
                                    End If
                                    
                                    If foundMatch Then
                                        ' Extract target currency (first 3 letters of DISPLAYED CcyPair)
                                        Dim targetCurrency As String
                                        targetCurrency = Left(displayedCcyPair, 3) ' e.g., USD from USDJPY
                                        
                                        ' Show FULL TSV position info
                                        Dim tsvFullPosition As String
                                        tsvFullPosition = "TSV " & tsvCurrencyPair & ": " & scnRowData(2) & " " & scnRowData(3) & _
                                                          " " & scnRowData(4) & " | " & scnRowData(5) & " " & scnRowData(6) & " " & scnRowData(7)
                                        
                                        ' Find which column contains the target currency
                                        Dim syntheticExposure As Double
                                        syntheticExposure = 0
                                        Dim amountValue As String
                                        Dim position As String
                                        
                                        ' Check Column 3 (index 2-3: position and currency)
                                        If scnRowData(3) = targetCurrency Then
                                            ' Found in Column 3
                                            amountValue = scnRowData(4)
                                            position = scnRowData(2)
                                            
                                        ' Check Column 6 (index 5-6: position and currency)
                                        ElseIf scnRowData(6) = targetCurrency Then
                                            ' Found in Column 6
                                            amountValue = scnRowData(7)
                                            position = scnRowData(5)
                                        Else
                                            ' Currency not found in either column
                                            debugInfo = debugInfo & " | " & tsvFullPosition & " | ERROR: " & targetCurrency & " not found"
                                            errorMessages.Add "Client: " & clientIDValue & ", CcyPair: " & displayedCcyPair & ", Error: " & targetCurrency & " not found"
                                            Exit For
                                        End If
                                        
                                        ' Add match details to debug info
                                        debugInfo = debugInfo & " | " & tsvFullPosition & " | Matched: " & position & " " & targetCurrency & " " & amountValue
                                        If wasFlippedForMatch Then
                                            debugInfo = debugInfo & " (flipped)"
                                        End If
                                        
                                        ' Parse amount (remove commas)
                                        amountValue = Replace(amountValue, ",", "")
                                        If IsNumeric(amountValue) Then
                                            syntheticExposure = CDbl(amountValue)
                                            
                                            ' Apply SHORT/LONG sign
                                            ' SIMPLE RULE: Use position from TSV directly
                                            If UCase(position) = "SHORT" Then
                                                syntheticExposure = -syntheticExposure
                                            ' LONG is positive, no change needed
                                            End If
                                            
                                            ' Convert to millions and round to 1 decimal
                                            syntheticExposure = Round(syntheticExposure / 1000000, 1)
                                            
                                            ' Store in Synth. Exp. column (E)
                                            totalData(i, 5) = syntheticExposure
                                            debugInfo = debugInfo & " | Synth: " & syntheticExposure
                                        Else
                                            debugInfo = debugInfo & " | ERROR: Invalid amount"
                                            errorMessages.Add "Client: " & clientIDValue & ", CcyPair: " & displayedCcyPair & ", Error: Invalid amount format: " & amountValue
                                        End If
                                        
                                        ' Add debug info to error column
                                        errorMessages.Add debugInfo
                                        
                                        Exit For
                                    End If
                                End If
                            Next k
                            
                            ' If no match found and there's SCN data, it means this currency pair isn't in Section D
                            If Not foundMatch Then
                                debugInfo = debugInfo & " | No match in Section D"
                                errorMessages.Add debugInfo
                            End If
                        Next i
                    ElseIf singleCurrencySuppressed Then
                        ' Single Currency Netting Suppressed - set all Synthetic Exposure to 0
                        For i = 1 To UBound(totalData, 1)
                            totalData(i, 5) = 0
                        Next i
                        errorMessages.Add "Client: " & clientIDValue & ", Info: Single currency netting suppressed, all synthetic exposures set to 0"
                    ElseIf scnSyntheticData.Count = 0 Then
                        ' No SCN data found at all
                        For i = 1 To UBound(totalData, 1)
                            Dim displayedPair As String
                            displayedPair = displayCcyPairs(i)
                            errorMessages.Add "Client: " & clientIDValue & ", CcyPair: " & displayedPair & ", Info: No SCN data in file"
                        Next i
                    End If
                    
WriteData:
                    ' Write main data block
                    ws.Cells(currentRow, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                    rowsWritten = UBound(totalData, 1)
                    
                    ' Write error messages to column R
                    If errorMessages.Count > 0 Then
                        For k = 1 To errorMessages.Count
                            ws.Cells(currentRow + k - 1, 18).Value = errorMessages(k)
                        Next k
                        rowsWritten = Application.Max(rowsWritten, errorMessages.Count)
                    End If
                    
                    ' Write FX rate data block to columns K & L (11 & 12)
                    If fxRates.Count > 0 Then
                        ReDim fxData(1 To fxRates.Count, 1 To 2)
                        i = 1
                        For Each key In fxRates.Keys
                            fxData(i, 1) = key
                            fxData(i, 2) = fxRates(key)
                            i = i + 1
                        Next key
                        ' Start writing FX data at the same row as the main block data
                        ws.Cells(currentRow, 11).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
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
                    
                    ' Clear old content in columns A:R if the new data is shorter
                    If currentRow + rowsWritten <= oldBlockEndRow Then
                        ws.Range(ws.Cells(currentRow + rowsWritten, 1), ws.Cells(oldBlockEndRow, 18)).ClearContents
                    End If
                    
                    ' Set currentRow to the row of the next ClientID found
                    currentRow = nextClientIDRow
                
                Else
                    ' ClientID is not empty, but we don't have a matching file (or it's an empty cell)
                    currentRow = currentRow + 1
                End If
            Loop
            
            ' Clean up after last client
            lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
            If currentRow + 2 <= lastRow Then
                ws.Range("A" & currentRow + 2 & ":B" & lastRow).ClearContents
            End If
            
            ' Add reference formulas
            ws.Cells(currentRow + 3, 1).Value = "Spot ref: Spot/Forward/Cash"
            ws.Cells(currentRow + 4, 1).Value = "USDJPY:"
            ws.Cells(currentRow + 5, 1).Value = "GBPJPY:"
            ws.Cells(currentRow + 7, 1).Value = "Spot ref: Structures"
            ws.Cells(currentRow + 8, 1).Value = "USDJPY:"
            ws.Cells(currentRow + 9, 1).Value = "GBPJPY:"
            
            With ws.Cells(currentRow + 4, 2)
                .Formula = "=L5" ' Now column L instead of K
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            With ws.Cells(currentRow + 5, 2)
                .Formula = "=L5*L4" ' Now column L instead of K
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            ' Final formatting
            ws.Range("A:R").Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    Range("D24").Value2 = "NOP" & Chr(10) & "($ mio)"
    MsgBox "Data update complete on sheet 'ContRisk3'!", vbInformation
    
End Sub
Sub ImportMultipleTSVAndStackDatav7() ' Updated version number
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
    Dim startSCNSynthetic As Boolean, endSCNSynthetic As Boolean
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
    Dim coverRatioString As String
    Dim exposureString As String
    Dim fxRateString As String
    
    ' Variables for the new in-place update logic
    Dim fileMap As Object
    Dim currentRow As Long
    Dim clientIDValue As String
    Dim rowsWritten As Long
    Dim nextClientIDRow As Long
    Dim oldBlockEndRow As Long

    ' Data structures for the blocks of data
    Dim dataCollection As Collection      ' For Total rows (columns O-P-Q)
    Dim cashTradeCollection As Collection ' For Cash Trade rows (columns S-T-U)
    Dim fxRates As Object
    Dim scnSyntheticData As Collection
    Dim clientID As String
    Dim coverRatio As Double
    Dim foundCover As Boolean
    Dim singleCurrencySuppressed As Boolean
    Dim errorMessages As Collection
    
    ' --- 1. Setup and File Selection ---
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set fileMap = CreateObject("Scripting.Dictionary")

    With fd
        .Title = "Select TSV Files"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = INITIAL_DIRECTORY
        .AllowMultiSelect = True

        If .Show = -1 Then
            ' Set the output worksheet to "ContRisk5"
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets("ContRisk5")
            On Error GoTo 0
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.Name = "ContRisk5"
                
                ' Set up headers for ContRisk5 with your specified column titles
                ws.Cells(1, 1).Value = "ClientID"                    ' Column A
                ws.Cells(1, 2).Value = ""                            ' Column B (blank)
                ws.Cells(1, 3).Value = ""                            ' Column C (blank)
                ws.Cells(1, 4).Value = ""                            ' Column D (blank)
                ws.Cells(1, 5).Value = ""                            ' Column E (blank)
                ws.Cells(1, 6).Value = ""                            ' Column F (blank)
                ws.Cells(1, 7).Value = ""                            ' Column G (blank)
                ws.Cells(1, 8).Value = "To Flip"                     ' Column H
                ws.Cells(1, 9).Value = "Divde by"                    ' Column I
                ws.Cells(1, 10).Value = "Mult by"                    ' Column J
                ws.Cells(1, 11).Value = ""                            ' Column K (blank)
                ws.Cells(1, 12).Value = "Currency"                    ' Column L
                ws.Cells(1, 13).Value = "Rate"                        ' Column M
                ws.Cells(1, 14).Value = ""                            ' Column N (blank)
                ws.Cells(1, 15).Value = "Exposure Ccy."               ' Column O
                ws.Cells(1, 16).Value = "Delta Amt"                   ' Column P
                ws.Cells(1, 17).Value = "Delta in Base Amt"           ' Column Q
                ws.Cells(1, 18).Value = ""                            ' Column R (blank)
                ws.Cells(1, 19).Value = "SCN. Exposure Ccy."          ' Column S
                ws.Cells(1, 20).Value = "SCN. Cash Delta Amt"         ' Column T
                ws.Cells(1, 21).Value = "SCN. Delta in Base Amt"      ' Column U
                ws.Cells(1, 22).Value = ""                            ' Column V (blank)
                ws.Cells(1, 23).Value = "Debug"                       ' Column W
            End If
            
            ' Pre-load the ToFlipTable from columns H, I, J
            Set ToFlipTable = New Collection
            k = 2 ' Start from row 2 to skip the header
            Dim wsCurrent As Worksheet
            Set wsCurrent = ThisWorkbook.Sheets(ws.Name)
            
            Do While Len(wsCurrent.Cells(k, 8).Value) > 0 And (Len(wsCurrent.Cells(k, 9).Value) > 0 Or Len(wsCurrent.Cells(k, 10).Value) > 0)
                flipEntry(0) = wsCurrent.Cells(k, 8).Value ' Currency pair from column H
                flipEntry(1) = wsCurrent.Cells(k, 9).Value ' Divisor currency from column I
                flipEntry(2) = wsCurrent.Cells(k, 10).Value ' Multiplier currency from column J
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

            'lastRow = ws.Cells(ws.rows.Count, "H").End(xlUp).Row
            lastRow = 28
            ' Clear columns
            If lastRow > 1 Then
                ws.Range("B4:F" & lastRow - 1).ClearContents
                ws.Range("L4:W" & lastRow - 1).ClearContents
            End If
            
            currentRow = 4 ' Start after the header
            
            Do While currentRow <= lastRow
                clientIDValue = Trim(ws.Cells(currentRow, 1).Value)

                ' Only update if a valid ClientID is found in Col A AND we have a file for it
                If Len(clientIDValue) > 0 And fileMap.Exists(clientIDValue) Then
                    filePath = fileMap(clientIDValue)
                    
                    ' Reset variables for processing the TSV file
                    Set dataCollection = New Collection          ' For Total rows
                    Set cashTradeCollection = New Collection     ' For Cash Trade rows
                    Set fxRates = CreateObject("Scripting.Dictionary")
                    Set scnSyntheticData = New Collection
                    Set errorMessages = New Collection
                    coverRatio = 0
                    foundCover = False
                    singleCurrencySuppressed = False
                    
                    ' --- Read and Process Data from TSV file ---
                    Open filePath For Input As #1
                    fileData = Input$(LOF(1), 1)
                    Close #1
                    
                    rows = Split(fileData, vbCrLf)
                    startSCNRates = False: endSCNRates = False
                    startRiskCashflow = False: endRiskCashflow = False
                    startSCNSynthetic = False: endSCNSynthetic = False
                    
                    For i = 0 To UBound(rows)
                        If Len(Trim(rows(i))) > 0 Then
                            rowData = Application.Trim(rows(i))
                            
                            ' Check for Single Currency Netting Suppressed
                            If InStr(rowData, "<*** SINGLE CURRENCY NETTING SUPPRESSED ***>") > 0 Then
                                singleCurrencySuppressed = True
                            End If
                            
                            ' Find Cover Ratio
                            If InStr(rowData, "Cover Ratio") > 0 Then
                                coverRatioString = Trim(Split(rowData, vbTab)(UBound(Split(rowData, vbTab))))
                                coverRatioString = Replace(coverRatioString, ",", "")
                                If IsNumeric(coverRatioString) Then coverRatio = CDbl(coverRatioString)
                                foundCover = True
                            End If
                            
                            ' Find FX Rates (Section B)
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
                            
                            ' Find SCN Synthetic Positions Data (Section D)
                            If rowData Like "D. SCN SYNTHETIC POSITIONS*" Then startSCNSynthetic = True
                            If rowData Like "E. SYNTHETIC POSITIONS*" Then endSCNSynthetic = True
                            If startSCNSynthetic And Not endSCNSynthetic Then
                                ' Skip divider lines and empty lines
                                rowData = Trim(rowData)
                                If Len(rowData) > 0 And Not (rowData Like "----------*" Or rowData Like "---*" Or rowData Like "----------------------------------------------------------------*") Then
                                    ' Parse the SCN Synthetic Positions data
                                    lineData = Split(rowData, vbTab)
                                    If UBound(lineData) >= 5 Then ' Need at least 6 columns (0-5)
                                        Dim scnRow(7) As String ' 8 columns (0-7) to hold all data
                                        For k = 0 To 7
                                            If k <= UBound(lineData) Then
                                                scnRow(k) = Trim(lineData(k))
                                            Else
                                                scnRow(k) = ""
                                            End If
                                        Next k
                                        
                                        ' Only add if first column has a currency pair (6 letters)
                                        If Len(scnRow(0)) >= 6 Then
                                            scnSyntheticData.Add scnRow
                                        End If
                                    End If
                                End If
                            End If
                            
                            ' Find Risk Cashflow Data (Section K)
                            If rowData Like "K. RISK CASHFLOW*" Then startRiskCashflow = True
                            If rowData Like "L. SEPARATED DIGITAL*" Then endRiskCashflow = True
                            
                            If startRiskCashflow And Not endRiskCashflow Then
                                ' Check for Total row
                                If rowData Like "Total*" Then
                                    lineData = Split(rowData, vbTab)
                                    If UBound(lineData) >= 6 Then
                                        Dim totalRow(2) As Variant
                                        totalRow(0) = lineData(2) ' CcyPair
                                        totalRow(1) = lineData(4) ' Risk Ccy
                                        exposureString = Trim(lineData(6))
                                        exposureString = Replace(exposureString, ",", "")
                                        exposureString = Replace(exposureString, "(", "-")
                                        exposureString = Replace(exposureString, ")", "")
                                        If IsNumeric(exposureString) Then totalRow(2) = CDbl(exposureString) Else totalRow(2) = 0
                                        dataCollection.Add totalRow
                                        
                                        ' Look ahead for Cash Trade row (next row)
                                        If i + 1 <= UBound(rows) Then
                                            Dim nextRow As String
                                            nextRow = Application.Trim(rows(i + 1))
                                            
                                            If InStr(nextRow, "Cash Trades") > 0 Then
                                                lineData = Split(nextRow, vbTab)
                                                If UBound(lineData) >= 5 Then
                                                    Dim cashRow(2) As Variant
                                                    ' Use same CcyPair and Risk Ccy from Total row
                                                    cashRow(0) = totalRow(0) ' CcyPair from Total
                                                    cashRow(1) = totalRow(1) ' Risk Ccy from Total
                                                    
                                                    ' Cash Trade amount
                                                    Dim cashExposure As String
                                                    cashExposure = Trim(lineData(6))
                                                    cashExposure = Replace(cashExposure, ",", "")
                                                    cashExposure = Replace(cashExposure, "(", "-")
                                                    cashExposure = Replace(cashExposure, ")", "")
                                                    If IsNumeric(cashExposure) Then cashRow(2) = CDbl(cashExposure) Else cashRow(2) = 0
                                                    
                                                    cashTradeCollection.Add cashRow
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                    
                    ' --- Prepare and Write Data In-Place ---

                    ' Determine size of the array for writing
                    If dataCollection.Count = 0 Then
                        ReDim totalData(1 To 1, 1 To 21) ' 21 columns (A-U)
                        
                        ' Just write basic client info if no data
                        totalData(1, 1) = clientIDValue
                        totalData(1, 2) = Int(coverRatio)
                        
                        ' Skip all the complex processing since there's no data
                        GoTo WriteData
                    Else
                        ReDim totalData(1 To dataCollection.Count, 1 To 21) ' A through U
                    End If

                    ' Store display versions of CcyPairs for matching
                    Dim displayCcyPairs() As String
                    ReDim displayCcyPairs(1 To dataCollection.Count)

                    For i = 1 To UBound(totalData, 1)
                        ' Populate ClientID and CoverRatio only for the first row
                        If i = 1 Then
                            totalData(i, 1) = clientIDValue ' Column A
                            totalData(i, 2) = Int(coverRatio) ' Column B
                        Else
                            totalData(i, 1) = ""
                            totalData(i, 2) = ""
                        End If
                        
                        ' Populate other columns from Total row data
                        Dim originalCcyPair As String
                        originalCcyPair = dataCollection(i)(0) ' Original CcyPair from Section K Total
                        
                        ' Apply DISPLAY FLIP logic to get what's shown in column C
                        Dim displayCcyPair As String
                        displayCcyPair = originalCcyPair ' Start with original
                        
                        ' Check if this pair needs to be flipped for display
                        For k = 1 To ToFlipTable.Count
                            currentFlipEntry = ToFlipTable(k)
                            
                            If originalCcyPair = currentFlipEntry(0) Then
                                ' This pair is in the ToFlipTable, so flip it for display
                                If Len(originalCcyPair) = 6 Then
                                    displayCcyPair = Right(originalCcyPair, 3) & Left(originalCcyPair, 3)
                                End If
                                Exit For
                            End If
                        Next k
                        
                        ' Store the display version
                        displayCcyPairs(i) = displayCcyPair
                        totalData(i, 3) = displayCcyPair ' Column C shows flipped version
                        
                        ' Columns O, P from Total row data
                        totalData(i, 15) = dataCollection(i)(1) ' Risk Ccy (column O)
                        totalData(i, 16) = dataCollection(i)(2) ' Exposure amount (column P)
                        
                        ' Columns S, T from Cash Trade data (if available)
                        If cashTradeCollection.Count >= i Then
                            totalData(i, 19) = cashTradeCollection(i)(1) ' SCN Exposure Ccy (column S) - same as Risk Ccy
                            totalData(i, 20) = cashTradeCollection(i)(2) ' SCN Cash Delta Amt (column T)
                        End If
                        
                        ' Apply manipulation logic for exposure calculation (for both P and T)
                        ccyPair = originalCcyPair ' Use original for calculation
                        
                        For k = 1 To ToFlipTable.Count
                            currentFlipEntry = ToFlipTable(k)
                            
                            If ccyPair = currentFlipEntry(0) Then
                                divRate = 1: mulRate = 1
                                divCurrency = currentFlipEntry(1)
                                mulCurrency = currentFlipEntry(2)
                                
                                ' Get FX rates
                                For Each key In fxRates.Keys
                                    If key = divCurrency Then divRate = fxRates(key)
                                    If key = mulCurrency Then mulRate = fxRates(key)
                                    If key = 1 Then
                                        mulRate = -1
                                        divRate = 1
                                    End If
                                    If key = -1 Then
                                        mulRate = 1
                                        divRate = 1
                                    End If
                                Next key
                                
                                ' Calculate Column Q (Delta in Base Amt) from Column P
                                If divRate <> 0 And mulRate <> 0 Then
                                    originalExposure = totalData(i, 16) ' Column P
                                    manipulatedExposure = (originalExposure / divRate) * mulRate * -1
                                    totalData(i, 17) = manipulatedExposure ' Column Q
                                    
                                    ' Calculate Column U (SCN Delta in Base Amt) from Column T
                                    If cashTradeCollection.Count >= i Then
                                        Dim scnExposure As Double
                                        scnExposure = totalData(i, 20) ' Column T
                                        If scnExposure <> 0 Then
                                            Dim scnManipulatedExposure As Double
                                            scnManipulatedExposure = (scnExposure / divRate) * mulRate * -1
                                            totalData(i, 21) = scnManipulatedExposure ' Column U
                                        End If
                                    End If
                                End If
                                
                                Exit For
                            End If
                        Next k
                        
                        ' Populate the "Delta (mio)" column (column D)
                        If Not IsEmpty(totalData(i, 17)) And IsNumeric(totalData(i, 17)) Then
                            valueForMio = CDbl(totalData(i, 17))
                        ElseIf IsNumeric(totalData(i, 16)) Then
                            valueForMio = CDbl(totalData(i, 16))
                        Else
                            valueForMio = 0
                        End If
                        totalData(i, 4) = Round(valueForMio / 1000000, 1)
                        
                        ' Initialize Synth. Exp. column (column E)
                        totalData(i, 5) = 0
                        
                        ' Column F (SOV) remains blank
                        totalData(i, 6) = ""
                        ' Column G remains blank
                        totalData(i, 7) = ""
                    Next i
                    
                    ' --- Process SCN Synthetic Positions Data for Synthetic Exposure (Column E) ---
                    If Not singleCurrencySuppressed And scnSyntheticData.Count > 0 Then
                        ' Process each currency pair from the main data
                        For i = 1 To UBound(totalData, 1)
                            ' Get the DISPLAYED CcyPair (what's shown in column C)
                            Dim displayedCcyPair As String
                            displayedCcyPair = displayCcyPairs(i)
                            
                            ' Debug info
                            Dim debugInfo As String
                            debugInfo = "Display: " & displayedCcyPair & ", Original: " & dataCollection(i)(0)
                            
                            ' Look for matching data in SCN Synthetic Positions
                            Dim foundMatch As Boolean
                            foundMatch = False
                            
                            For k = 1 To scnSyntheticData.Count
                                Dim scnRowData() As String
                                scnRowData = scnSyntheticData(k)
                                
                                ' Check if we have valid data
                                If UBound(scnRowData) >= 6 Then
                                    Dim tsvCurrencyPair As String
                                    tsvCurrencyPair = scnRowData(0)
                                    
                                    ' Try to match against DISPLAYED pair
                                    Dim tsvFlippedPair As String
                                    If Len(tsvCurrencyPair) = 6 Then
                                        tsvFlippedPair = Right(tsvCurrencyPair, 3) & Left(tsvCurrencyPair, 3)
                                    Else
                                        tsvFlippedPair = ""
                                    End If
                                    
                                    ' Check for match
                                    If tsvCurrencyPair = displayedCcyPair Or tsvFlippedPair = displayedCcyPair Then
                                        foundMatch = True
                                        
                                        ' Extract target currency (first 3 letters of DISPLAYED CcyPair)
                                        Dim targetCurrency As String
                                        targetCurrency = Left(displayedCcyPair, 3)
                                        
                                        ' Find which column contains the target currency
                                        Dim syntheticExposure As Double
                                        syntheticExposure = 0
                                        Dim amountValue As String
                                        Dim position As String
                                        
                                        ' Check Column 3 (index 2-3: position and currency)
                                        If scnRowData(3) = targetCurrency Then
                                            amountValue = scnRowData(4)
                                            position = scnRowData(2)
                                            
                                        ' Check Column 6 (index 5-6: position and currency)
                                        ElseIf scnRowData(6) = targetCurrency Then
                                            amountValue = scnRowData(7)
                                            position = scnRowData(5)
                                        Else
                                            debugInfo = debugInfo & " | ERROR: " & targetCurrency & " not found"
                                            errorMessages.Add "Client: " & clientIDValue & ", CcyPair: " & displayedCcyPair & ", Error: " & targetCurrency & " not found"
                                            Exit For
                                        End If
                                        
                                        ' Parse amount
                                        amountValue = Replace(amountValue, ",", "")
                                        If IsNumeric(amountValue) Then
                                            syntheticExposure = CDbl(amountValue)
                                            
                                            ' Apply SHORT/LONG sign
                                            If UCase(position) = "SHORT" Then
                                                syntheticExposure = -syntheticExposure
                                            End If
                                            
                                            ' Convert to millions and round to 1 decimal
                                            syntheticExposure = Round(syntheticExposure / 1000000, 1)
                                            
                                            ' Store in Synth. Exp. column (E)
                                            totalData(i, 5) = syntheticExposure
                                        Else
                                            errorMessages.Add "Client: " & clientIDValue & ", CcyPair: " & displayedCcyPair & ", Error: Invalid amount format: " & amountValue
                                        End If
                                        
                                        Exit For
                                    End If
                                End If
                            Next k
                            
                            If Not foundMatch Then
                                debugInfo = debugInfo & " | No match in Section D"
                                errorMessages.Add debugInfo
                            End If
                        Next i
                    End If
                    
                    ' --- SPECIAL HANDLING FOR CLIENTS WITH SINGLE CURRENCY NETTING SUPPRESSED ---
                    If singleCurrencySuppressed Then
                        ' Overwrite Column E with values from Column U or T
                        For i = 1 To UBound(totalData, 1)
                            If Not IsEmpty(totalData(i, 21)) And IsNumeric(totalData(i, 21)) Then
                                valueForMio = CDbl(totalData(i, 21))
                            ElseIf IsNumeric(totalData(i, 20)) Then
                                valueForMio = CDbl(totalData(i, 20))
                            Else
                                valueForMio = 0
                            End If
                            totalData(i, 5) = Round(valueForMio / 1000000, 1)
                        Next i
                        
                        ' Get the number of data rows (currency pairs)
                        Dim dataRows As Long
                        dataRows = UBound(totalData, 1)
                        
                        ' Create a new array with required rows (at least 4)
                        Dim targetRows As Long
                        targetRows = Application.Max(dataRows, 4)
                        
                        Dim newData() As Variant
                        ReDim newData(1 To targetRows, 1 To 21)
                        
                        ' Copy existing data
                        For i = 1 To dataRows
                            For k = 1 To 21
                                newData(i, k) = totalData(i, k)
                            Next k
                        Next i
                        
                        ' Add messages
                        ' Row 2: "SCN suppressed" in column A
                        newData(2, 1) = "SCN suppressed"
                        
                        ' Row 3: second message in column A
                        newData(3, 1) = "Synth Pos col has cash delta in base ccy"
                        
                        ' Ensure row 4 is blank (it already is in new array)
                        
                        ' Replace totalData with newData
                        totalData = newData
                        
                        ' Add debug info
                        errorMessages.Add "Client: " & clientIDValue & ", Info: Single currency netting suppressed - Messages added in rows 2-3"
                    End If
WriteData:
                    ' Write main data block (columns A through U)
                    ws.Cells(currentRow, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
                    rowsWritten = UBound(totalData, 1)
                    
                    ' Write error messages to column W
                    If errorMessages.Count > 0 Then
                        For k = 1 To errorMessages.Count
                            ws.Cells(currentRow + k - 1, 23).Value = errorMessages(k) ' Column W = 23
                        Next k
                        rowsWritten = Application.Max(rowsWritten, errorMessages.Count)
                    End If
                    
                    ' Write FX rate data block to columns L & M (12 & 13)
                    If fxRates.Count > 0 Then
                        ReDim fxData(1 To fxRates.Count, 1 To 2)
                        i = 1
                        For Each key In fxRates.Keys
                            fxData(i, 1) = key
                            fxData(i, 2) = fxRates(key)
                            i = i + 1
                        Next key
                        ws.Cells(currentRow, 12).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
                        rowsWritten = Application.Max(rowsWritten, UBound(fxData, 1))
                    End If
                    
                    ' --- Clean up old data and set cursor for next loop ---
                    
                    ' Find the row where the next ClientID is listed
                    nextClientIDRow = currentRow + 1
                    Do While nextClientIDRow <= lastRow And Len(Trim(ws.Cells(nextClientIDRow, 1).Value)) = 0
                        nextClientIDRow = nextClientIDRow + 1
                    Loop
                    
                    ' Calculate the area of old data to clear
                    oldBlockEndRow = nextClientIDRow - 1
                    
                    ' Clear old content if new data is shorter
                    If currentRow + rowsWritten <= oldBlockEndRow Then
                        ws.Range(ws.Cells(currentRow + rowsWritten, 1), ws.Cells(oldBlockEndRow, 23)).ClearContents
                    End If
                    
                    ' Set currentRow to the row of the next ClientID found
                    currentRow = nextClientIDRow
                
                Else
                    ' ClientID is not empty, but we don't have a matching file
                    currentRow = currentRow + 1
                End If
            Loop
            
            ' Clean up after last client
            lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
            If currentRow + 2 <= lastRow Then
                ws.Range("A" & currentRow + 2 & ":B" & lastRow).ClearContents
            End If
            
            ' Add reference formulas
            ws.Cells(currentRow + 3, 1).Value = "Spot ref: Spot/Forward/Cash"
            ws.Cells(currentRow + 4, 1).Value = "USDJPY:"
            ws.Cells(currentRow + 5, 1).Value = "GBPJPY:"
            ws.Cells(currentRow + 7, 1).Value = "Spot ref: Structures"
            ws.Cells(currentRow + 8, 1).Value = "USDJPY:"
            ws.Cells(currentRow + 9, 1).Value = "GBPJPY:"
            
            ' Update formulas to use column M
            With ws.Cells(currentRow + 4, 2)
                .Formula = "=M5"
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            With ws.Cells(currentRow + 5, 2)
                .Formula = "=M5*M4"
                .numberFormat = "0.00"
                .HorizontalAlignment = xlLeft
            End With
            
            ' Final formatting
            ws.Range("A:W").Columns.AutoFit
            
        Else
            MsgBox "No file selected.", vbExclamation
        End If
    End With
    
    ' Update this reference
    Range("D24").Value2 = "NOP" & Chr(10) & "($ mio)"
    
    MsgBox "Data update complete on sheet 'ContRisk5'!", vbInformation
    
End Sub
