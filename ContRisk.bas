Sub ImportTSVRisk()
    Dim wsRisk As Worksheet
    Dim filePath As String
    Dim fileData As String
    Dim rows() As String
    Dim rowData As String
    Dim tempData As Variant
    Dim lineData() As String
    Dim startSection As Boolean, endSection As Boolean
    Dim i As Long, j As Long
    Dim exposureRiskCcy As Double, exposureUsd As Double
    Dim rowCount As Long
    Dim dataCollection As Collection
    Dim finalData() As Variant
    
    ' Reference the Risk worksheet
    'ThisWorkbook.Activate
    Set wsRisk = Worksheets("Risk")
    
    ' Open file dialog to select the TSV file from C:\Downloads\
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
            wsRisk.Range("B1").Value = filePath  ' Store file path in cell B1
        Else
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    If wsRisk.AutoFilterMode Then wsRisk.AutoFilterMode = False
    Dim lastRow As Long
    
    lastRow = wsRisk.Cells(wsRisk.rows.Count, "A").End(xlUp).Row
    If lastRow >= 5 Then
    ' Clear content from row 5 onwards
    wsRisk.rows("5:" & lastRow).ClearContents
    End If
    ' Define headers
    Dim headers As Variant
    headers = Array("Total", "Orig/Recast", "CcyPair", "Date", "RiskCCy", "MTM", _
                    "Exposure (RiskCCy)", "Exposure (USD)", "GammaAddon", "VegaAddon", _
                    "Gamma+VegaAddon", "Basic Cover", "Total Cover")
    
    ' Open the TSV file and read all data into memory
    Open filePath For Input As #1
    fileData = Input$(LOF(1), 1)
    Close #1
    
    ' Split file data into rows
    rows = Split(fileData, vbCrLf)
    
    ' Initialize variables for data processing
    startSection = False
    endSection = False
    rowCount = 0
    Set dataCollection = New Collection
    
    ' Add headers to the collection as the first row
    dataCollection.Add headers
    
    ' Loop through rows to find the section between "K. RISK CASHFLOW" and "L. SEPARATED DIGITAL"
    For i = 0 To UBound(rows)
        rowData = Application.Trim(rows(i))  ' Trim extra spaces
        
        ' Check for start of "K. RISK CASHFLOW" section
        If rowData Like "K. RISK CASHFLOW*" Then
            startSection = True
            i = i + 2 ' Skip the next two rows: the "-------" and the immediate row after K. RISK CASHFLOW
            GoTo NextRow
        End If
        
        ' Check for end of section at "L. SEPARATED DIGITAL"
        If rowData Like "L. SEPARATED DIGITAL*" Then
            endSection = True
            Exit For
        End If
        
        ' If within section and row starts with "Total," store the row
        If startSection And Not endSection Then
            If rowData Like "Total*" Then
                ' Split by tab and prepare the row for "Original" and "Recast"
                lineData = Split(rowData, vbTab)
                
                ' Copy data for "Original" row
                Dim originalRow(12) As Variant
                For j = 0 To UBound(lineData)
                    originalRow(j) = lineData(j)
                Next j
                originalRow(1) = "Original"  ' Set Orig/Recast as "Original"
                
                ' Set Exposure (RiskCCy) and Exposure (USD) for original row
                exposureRiskCcy = CDbl(lineData(6))  ' Assuming Exposure (RiskCCy) is at index 6 in the array
                exposureUsd = CDbl(lineData(7))      ' Assuming Exposure (USD) is at index 7 in the array
                
                ' Add the original row to the collection
                dataCollection.Add originalRow
                
                ' Prepare "Recast" row with modifications
                Dim recastRow(12) As Variant
                For j = 0 To UBound(lineData)
                    recastRow(j) = lineData(j)
                Next j
                recastRow(1) = "Recast"  ' Set Orig/Recast as "Recast"
                
                ' Modify Exposure (USD) in "Recast" row if Exposure (RiskCCy) is positive
                If exposureRiskCcy > 0 Then
                    recastRow(7) = -exposureUsd  ' Set Exposure (USD) in the recast row to negative
                End If
                
                ' Add the recast row to the collection
                dataCollection.Add recastRow
            End If
        End If
NextRow:
    Next i
    
    ' Transfer data from collection to array for final output
    rowCount = dataCollection.Count
    ReDim finalData(1 To rowCount, 1 To UBound(headers) + 1)
    
    ' Copy each item in the collection to the final data array
    For i = 1 To rowCount
        tempData = dataCollection(i)
        For j = 1 To UBound(headers) + 1
            finalData(i, j) = tempData(j - 1)
        Next j
    Next i
    
    ' Write the entire data array to the worksheet in one operation
    wsRisk.Cells(5, 1).Resize(UBound(finalData, 1), UBound(finalData, 2)).Value = finalData
    
    ' Apply filter and sort "Recast" data by "Total Cover" in descending order
    With wsRisk
        .Range("A5").CurrentRegion.AutoFilter Field:=2, Criteria1:="Recast" ' Filter "Orig/Recast" column for "Recast"
        .Range("A5").CurrentRegion.Sort Key1:=.Range("M5"), Order1:=xlDescending, Header:=xlYes ' Sort by "Total Cover"
    End With
End Sub

Sub CopyVisibleData()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim rngSource As Range, rngVisible As Range, rngArea As Range
    Dim lastRow As Long, rowCount As Long
    Dim destRange As Range
    
    ' Set source and destination sheets
    Set wsSource = ThisWorkbook.Sheets("Risk")
    Set wsDest = ThisWorkbook.Sheets("Compare")
    
    ' Determine last row in source
    lastRow = wsSource.Cells(wsSource.rows.Count, "A").End(xlUp).Row
    If lastRow < 6 Then Exit Sub ' No data to copy
    
    ' Define source range (excluding headers)
    Set rngSource = wsSource.Range("A5:M" & lastRow)
    
    ' Get visible cells only
    On Error Resume Next
    Set rngVisible = rngSource.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If rngVisible Is Nothing Then Exit Sub ' No visible data to copy
    
    ' Count rows to insert (excluding headers)r
    rowCount = 0
    For Each rngArea In rngVisible.Areas
        rowCount = rowCount + rngArea.rows.Count
    Next rngArea
    
    ' Insert rows at the top of "Compare"
    wsDest.rows("1:" & rowCount + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Paste values into columns B:N
    Set destRange = wsDest.Range("B1")
    rngVisible.Copy
    destRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Insert today's date in column A
    wsDest.Range("A2:A" & rowCount).Value = Date
    wsDest.Columns("A").NumberFormat = "dd mmm yy"
    
    ' Make header row bold
    wsDest.rows(1).Font.Bold = True
End Sub


