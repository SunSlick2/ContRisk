Sub ImportTSVAndWriteData()

    Dim ws As Worksheet
    Dim filePath As String
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long, j As Long
    
    ' Boolean flags for section processing
    Dim startRiskCashflow As Boolean, endRiskCashflow As Boolean
    Dim startSCNRates As Boolean, endSCNRates As Boolean
    
    ' Data structures for the blocks of data
    Dim dataCollection As Collection
    Set dataCollection = New Collection
    Dim fxRates As Object ' Dictionary for FX rates
    Set fxRates = CreateObject("Scripting.Dictionary")
    
    ' Variables for new data points
    Dim clientID As String
    Dim coverRatio As Double
    Dim foundClient As Boolean, foundCover As Boolean
    
    ' --- 1. File Selection ---
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select a TSV File"
        .Filters.Clear
        .Filters.Add "TSV Files", "*.tsv"
        .InitialFileName = "C:\Users\abc\Downloads\MMDump\"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Reference the active worksheet
    Set ws = ActiveSheet
    
    ' --- 2. Read and Process the Data ---
    
    Open filePath For Input As #1
    fileData = Input$(LOF(1), 1)
    Close #1
    
    rows = Split(fileData, vbCrLf)
    
    ' Loop through rows to find all required sections and data
    For i = 0 To UBound(rows)
        rowData = Application.Trim(rows(i))
        
        ' Capture Client ID (close to top of file)
        If Not foundClient Then
            If InStr(rowData, "Client:") > 0 Then
                Dim parts() As String
                parts = Split(rowData, "Client:")
                If UBound(parts) > 0 Then
                    clientID = Trim(parts(1))
                    foundClient = True
                End If
            End If
        End If
        
        ' Capture Cover Ratio (close to end of file)
        If InStr(rowData, "Cover Ratio") > 0 Then
            Dim coverParts() As String
            coverParts = Split(rowData, vbTab)
            If UBound(coverParts) > 0 Then
                coverRatio = CDbl(Trim(coverParts(UBound(coverParts))))
                foundCover = True
            End If
        End If
        
        ' Check for start/end of SCN RATES section
        If UCase(rowData) Like "B. SCN RATES*" Then
            startSCNRates = True
            i = i + 1 ' Skip the header row
        ElseIf UCase(rowData) Like "C. SCN BREAKDOWN*" Then
            endSCNRates = True
        End If
        
        ' If inside SCN RATES section, capture FX rates
        If startSCNRates And Not endSCNRates Then
            If rowData Like "FX.Rate.*.Spot" Then
                Dim lineData As Variant
                lineData = Split(rowData, vbTab)
                
                Dim ccy As String
                ' Extract currency by removing "FX.Rate." and ".Spot"
                ccy = Mid(lineData(0), Len("FX.Rate.") + 1, Len(lineData(0)) - Len("FX.Rate.") - Len(".Spot"))
                If IsNumeric(lineData(1)) Then
                    fxRates.Add ccy, CDbl(lineData(1))
                End If
            End If
        End If
        
        ' Check for start/end of Risk Cashflow section
        If rowData Like "K. RISK CASHFLOW*" Then
            startRiskCashflow = True
            i = i + 2 ' Skip the next two rows
        ElseIf rowData Like "L. SEPARATED DIGITAL*" Then
            endRiskCashflow = True
        End If
        
        ' If inside Risk Cashflow section, capture "Total" rows
        If startRiskCashflow And Not endRiskCashflow Then
            If rowData Like "Total*" Then
                lineData = Split(rowData, vbTab)
                
                Dim totalRow(2) As Variant
                totalRow(0) = lineData(2) ' CcyPair
                totalRow(1) = lineData(4) ' RiskCCy
                totalRow(2) = lineData(6) ' Exposure (RiskCCy)
                
                dataCollection.Add totalRow
            End If
        End If
    Next i
    
    ' --- 3. Write Data to Worksheet ---
    
    ws.UsedRange.ClearContents
    
    Dim lastRow As Long
    lastRow = 1
    
    ' Write Client and Cover Ratio data
    ws.Cells(lastRow, 1).Value = "Client ID"
    ws.Cells(lastRow, 2).Value = clientID
    ws.Cells(lastRow + 1, 1).Value = "Cover Ratio"
    ws.Cells(lastRow + 1, 2).Value = coverRatio
    lastRow = lastRow + 3 ' Leave a gap
    
    ' Write the "Total" row data block
    If dataCollection.Count > 0 Then
        Dim totalData() As Variant
        Dim totalHeaders As Variant
        ReDim totalData(1 To dataCollection.Count, 1 To 3)
        For i = 1 To dataCollection.Count
            For j = 0 To 2
                totalData(i, j + 1) = dataCollection(i)(j)
            Next j
        Next i
        
        totalHeaders = Array("CcyPair", "RiskCCy", "Exposure (RiskCCy)")
        ws.Cells(lastRow, 1).Resize(1, UBound(totalHeaders) + 1).Value = totalHeaders
        ws.Cells(lastRow + 1, 1).Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
    End If
    
    ' Write the FX rate data block
    If fxRates.Count > 0 Then
        Dim fxData() As Variant
        Dim fxHeaders As Variant
        ReDim fxData(1 To fxRates.Count, 1 To 2)
        i = 1
        For Each key In fxRates.Keys
            fxData(i, 1) = key
            fxData(i, 2) = fxRates(key)
            i = i + 1
        Next key
        
        fxHeaders = Array("Currency", "Mid Spot Rate")
        ws.Cells(lastRow, 1).Resize(1, UBound(fxHeaders) + 1).Value = fxHeaders
        ws.Cells(lastRow + 1, 1).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
    End If
    
    ' Optional: Formatting
    ws.UsedRange.Columns.AutoFit
    ws.Rows(1).Font.Bold = True
    ws.Rows(3).Font.Bold = True
    If ws.Cells(ws.Rows.Count, "A").End(xlUp).Row > 1 Then
        ws.Rows(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - fxRates.Count).Font.Bold = True
    End If

    MsgBox "Data import and writing complete!", vbInformation
    
End Sub