Sub CreatePivotTable()
    Dim wb As Workbook
    Dim wsData As Worksheet, wsFiltered As Worksheet, wsPivot As Worksheet
    Dim rngFiltered As Range
    Dim lastRow As Long
    Dim pCache As PivotCache
    Dim pTable As PivotTable

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("RHACS_Vulnerability_Report_Work")

    ' Delete old sheets if they exist
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Filtered").Delete
    wb.Sheets("FilteredPivot").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create Filtered sheet
    Set wsFiltered = wb.Sheets.Add(After:=wsData)
    wsFiltered.Name = "Filtered"

    ' Copy header row
    wsData.Rows(1).Copy Destination:=wsFiltered.Rows(1)
    wsFiltered.Cells(1, 13).Value = "CVE_Count" ' Helper column

    ' Filter and copy data
    Dim i As Long, j As Long, rowOut As Long
    rowOut = 2
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row

    For i = 2 To lastRow
        Dim ns As String, sev As String, cve As String
        ns = LCase(Trim(wsData.Cells(i, 2).Value))
        sev = UCase(Trim(wsData.Cells(i, 9).Value))
        cve = Trim(wsData.Cells(i, 6).Value)

        If ( _
            ns Like "openshift-*" Or ns Like "kube-*" Or _
            ns Like "rhacs-operator*" Or ns Like "open-cluster-management*" Or _
            ns Like "cert-manager*" Or _
            ns = "stackrox" Or ns = "multicluster-engine" Or _
            ns = "aap" Or ns = "hive" Or ns = "nvidia-gpu-operator") _
            And (sev = "CRITICAL" Or sev = "IMPORTANT") _
            And LCase(Left(cve, 4)) = "cve-" Then

            For j = 1 To 12
                wsFiltered.Cells(rowOut, j).Value = wsData.Cells(i, j).Value
            Next j
            wsFiltered.Cells(rowOut, 13).Value = 1 ' Add count
            rowOut = rowOut + 1
        End If
    Next i

    If rowOut = 2 Then
        MsgBox "No matching CVEs found.", vbExclamation
        Exit Sub
    End If

    ' Create Pivot Table sheet
    Set wsPivot = wb.Sheets.Add(After:=wsFiltered)
    wsPivot.Name = "FilteredPivot"

    Set rngFiltered = wsFiltered.Range("A1").CurrentRegion
    Set pCache = wb.PivotCaches.Create(xlDatabase, rngFiltered.Address(, , , True))

    Set pTable = pCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), TableName:="FilteredCVEPivot")

    ' Configure pivot fields
    With pTable
        .ClearAllFilters

        With .PivotFields("CVE")
            .Orientation = xlRowField
            .Position = 1
        End With

        With .PivotFields("Fixable")
            .Orientation = xlRowField
            .Position = 2
        End With

        With .PivotFields("Reference")
            .Orientation = xlRowField
            .Position = 3
        End With

        With .PivotFields("Component")
            .Orientation = xlRowField
            .Position = 4
        End With

        .AddDataField .PivotFields("CVE_Count"), "Count of CVE", xlSum
        .RowAxisLayout xlCompactRow
        .RepeatAllLabels xlRepeatLabels
        .RowGrand = True
        .ColumnGrand = False
    End With

    ' Rename Row Labels
    wsPivot.Range("A3").Value = "CVE/Fixable/Reference/Component"
    wsPivot.Columns("A:B").AutoFit

    ' === ? Insert Summary Rows ===
    Dim lastPivotRow As Long
    lastPivotRow = wsPivot.Cells(wsPivot.Rows.Count, "A").End(xlUp).row

    ' Count total CVEs and unique CVE IDs
    Dim totalCVEs As Long: totalCVEs = 0
    Dim uniqueCVEs As Long: uniqueCVEs = 0

    Dim row As Long
    For row = 4 To lastPivotRow
        If Left(wsPivot.Cells(row, 1).Value, 4) = "CVE-" Then
            totalCVEs = totalCVEs + wsPivot.Cells(row, 2).Value
            uniqueCVEs = uniqueCVEs + 1
        End If
    Next row

    ' Insert Unique CVEs count row
    wsPivot.Cells(lastPivotRow + 1, 1).Value = "Unique CVEs"
    wsPivot.Cells(lastPivotRow + 1, 2).Value = uniqueCVEs
    wsPivot.Cells(lastPivotRow + 1, 1).Font.Bold = True
    wsPivot.Cells(lastPivotRow + 1, 2).Font.Bold = True

    ' Insert Grand Total row
    wsPivot.Cells(lastPivotRow + 2, 1).Value = "Grand Total"
    wsPivot.Cells(lastPivotRow + 2, 2).Value = totalCVEs
    wsPivot.Cells(lastPivotRow + 2, 1).Font.Bold = True
    wsPivot.Cells(lastPivotRow + 2, 2).Font.Bold = True

    MsgBox "? Pivot created with Unique CVE count and correct Grand Total!", vbInformation
End Sub
