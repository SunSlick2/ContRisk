Sub ImportTSVAndWriteData()

    Dim ws As Worksheet
    Dim filePath As String
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim i As Long
    Dim startRiskCashflow As Boolean, endRiskCashflow As Boolean
    Dim startSCNRates As Boolean, endSCNRates As Boolean
    
    ' Data structures for the two blocks of data
    Dim dataCollection As Collection
    Set dataCollection = New Collection
    Dim fxRates As Object ' Use a dictionary for FX rates
    Set fxRates = CreateObject("Scripting.Dictionary")
    
    ' --- 1. File Selection ---
    
    ' Open file dialog to select the TSV file
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
    
    ' Open the TSV file and read all data into memory
    Open filePath For Input As #1
    fileData = Input$(LOF(1), 1)
    Close #1
    
    ' Split file data into rows
    rows = Split(fileData, vbCrLf)
    
    ' Loop through rows to find the sections
    For i = 0 To UBound(rows)
        rowData = Application.Trim(rows(i))
        
        ' Check for start/end of SCNRates section
        If rowData Like "B. SCNRates*" Then
            startSCNRates = True
            i = i + 1 ' Skip the header row
        ElseIf rowData Like "C. SCN Breakdown*" Then
            endSCNRates = True
        End If
        
        ' If inside SCNRates section, capture FX rates
        If startSCNRates And Not endSCNRates Then
            If rowData Like "FX.Rate.*" Then
                Dim lineData As Variant
                lineData = Split(rowData, vbTab)
                
                Dim ccy As String
                ccy = Right(lineData(0), Len(lineData(0)) - Len("FX.Rate.")) ' Extract currency name
                If IsNumeric(lineData(1)) Then
                    fxRates.Add ccy, CDbl(lineData(1)) ' Add to dictionary
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
                
                ' Capture CcyPair, RiskCCy, and Exposure (RiskCCy)
                Dim totalRow(2) As Variant
                totalRow(0) = lineData(2) ' CcyPair
                totalRow(1) = lineData(4) ' RiskCCy
                totalRow(2) = lineData(6) ' Exposure (RiskCCy)
                
                ' Add the new row to the collection
                dataCollection.Add totalRow
            End If
        End If
    Next i
    
    ' --- 3. Write Data to Worksheet ---
    
    ' Clear existing content from the sheet for a fresh import
    ws.UsedRange.ClearContents
    
    ' Write the "Total" row data block
    If dataCollection.Count > 0 Then
        Dim totalData() As Variant
        Dim totalHeaders As Variant
        Dim j As Long
        
        ReDim totalData(1 To dataCollection.Count, 1 To 3)
        For i = 1 To dataCollection.Count
            For j = 0 To 2
                totalData(i, j + 1) = dataCollection(i)(j)
            Next j
        Next i
        
        totalHeaders = Array("CcyPair", "RiskCCy", "Exposure (RiskCCy)")
        ws.Range("A1").Resize(1, UBound(totalHeaders) + 1).Value = totalHeaders
        ws.Range("A2").Resize(UBound(totalData, 1), UBound(totalData, 2)).Value = totalData
    End If
    
    ' Write the FX rate data block
    If fxRates.Count > 0 Then
        Dim fxData() As Variant
        Dim fxHeaders As Variant
        Dim startRowFX As Long
        
        ReDim fxData(1 To fxRates.Count, 1 To 2)
        i = 1
        For Each key In fxRates.Keys
            fxData(i, 1) = key
            fxData(i, 2) = fxRates(key)
            i = i + 1
        Next key
        
        fxHeaders = Array("Currency", "Mid Spot Rate")
        startRowFX = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
        
        ws.Cells(startRowFX, 1).Resize(1, UBound(fxHeaders) + 1).Value = fxHeaders
        ws.Cells(startRowFX + 1, 1).Resize(UBound(fxData, 1), UBound(fxData, 2)).Value = fxData
    End If
    
    ' Optional: Format the output
    ws.Columns("A:C").AutoFit
    ws.Rows(1).Font.Bold = True
    If startRowFX > 0 Then ws.Rows(startRowFX).Font.Bold = True

    MsgBox "Data import and writing complete!", vbInformation
    
End Sub