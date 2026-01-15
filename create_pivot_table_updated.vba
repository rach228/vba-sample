Sub CreatePivotTable()

    Dim wb As Workbook
    Dim wsData As Worksheet, wsFiltered As Worksheet, wsPivot As Worksheet
    Dim rngFiltered As Range
    Dim lastRow As Long
    Dim pCache As PivotCache
    Dim pTable As PivotTable

    Set wb = ThisWorkbook
    ' === CHANGE THIS TO YOUR SOURCE WORKSHEET NAME ===
    Set wsData = wb.Sheets("RHACS_Vulnerability_Report_Work")

    ' === Delete old sheets ===
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Filtered").Delete
    wb.Sheets("FilteredPivot").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' === Create filtered sheet ===
    Set wsFiltered = wb.Sheets.Add(After:=wsData)
    wsFiltered.Name = "Filtered"

    wsData.Rows(1).Copy Destination:=wsFiltered.Rows(1)
    wsFiltered.Cells(1, 13).Value = "CVE_Count"

    Dim i As Long, j As Long, rowOut As Long
    rowOut = 2
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' === NAMESPACE FILTER LIST (EDIT THIS) ===
    Dim nsList As Variant
    nsList = Array( _
        "openshift-*", _
        "kube-*", _
        "rhacs-operator*", _
        "open-cluster-management*", _
        "cert-manager*", _
        "ack-system", _
        "aws-load-balancer-operator", _
        "cert-utils-operator", _
        "costmanagement-metrics-operator", _
        "external-dns-operator", _
        "metallb-system", _
        "mtr", _
        "stackrox", _
        "multicluster-engine", _
        "multicluster-global-hub", _
        "node-observability-operator", _
        "aap", _
        "hive", _
        "redhat-ods-operator", _
        "rhdh-operator", _
        "service-telemetry", _
        "submariner-operator" _
    )

    ' === FILTER LOOP ===
    For i = 2 To lastRow
        
        Dim ns As String, sev As String, cve As String
        ns = LCase(Trim(wsData.Cells(i, 2).Value))
        sev = UCase(Trim(wsData.Cells(i, 9).Value))
        cve = Trim(wsData.Cells(i, 6).Value)

        If NamespaceMatches(ns, nsList) _
           And (sev = "CRITICAL" Or sev = "IMPORTANT") _
           And LCase(Left(cve, 4)) = "cve-" Then

            For j = 1 To 12
                wsFiltered.Cells(rowOut, j).Value = wsData.Cells(i, j).Value
            Next j

            wsFiltered.Cells(rowOut, 13).Value = 1
            rowOut = rowOut + 1
        End If
    Next i

    If rowOut = 2 Then
        MsgBox "No matching CVEs found.", vbExclamation
        Exit Sub
    End If

    ' === Create Pivot Table Sheet ===
    Set wsPivot = wb.Sheets.Add(After:=wsFiltered)
    wsPivot.Name = "FilteredPivot"

    Set rngFiltered = wsFiltered.Range("A1").CurrentRegion
    Set pCache = wb.PivotCaches.Create(xlDatabase, rngFiltered.Address(, , , True))

    Set pTable = pCache.CreatePivotTable(wsPivot.Range("A3"), "FilteredCVEPivot")

    With pTable
        .ClearAllFilters

        .PivotFields("CVE").Orientation = xlRowField
        .PivotFields("Fixable").Orientation = xlRowField
        .PivotFields("Reference").Orientation = xlRowField
        .PivotFields("Component").Orientation = xlRowField

        .AddDataField .PivotFields("CVE_Count"), "Count of CVE", xlSum
        .RowAxisLayout xlCompactRow
        .RepeatAllLabels xlRepeatLabels
    End With

    wsPivot.Range("A3").Value = "CVE/Fixable/Reference/Component"
    wsPivot.Columns("A:B").AutoFit

    ' === Additional Summary Rows ===
    Dim lastPivotRow As Long, row As Long
    Dim totalCVEs As Long: totalCVEs = 0
    Dim uniqueCVEs As Long: uniqueCVEs = 0

    lastPivotRow = wsPivot.Cells(wsPivot.Rows.Count, "A").End(xlUp).Row

    For row = 4 To lastPivotRow
        If Left(wsPivot.Cells(row, 1).Value, 4) = "CVE-" Then
            totalCVEs = totalCVEs + wsPivot.Cells(row, 2).Value
            uniqueCVEs = uniqueCVEs + 1
        End If
    Next row

    wsPivot.Cells(lastPivotRow + 1, 1).Value = "Unique CVEs"
    wsPivot.Cells(lastPivotRow + 1, 2).Value = uniqueCVEs
    wsPivot.Cells(lastPivotRow + 2, 1).Value = "Grand Total"
    wsPivot.Cells(lastPivotRow + 2, 2).Value = totalCVEs

    wsPivot.Range("A" & lastPivotRow + 1 & ":B" & lastPivotRow + 2).Font.Bold = True

    MsgBox "Pivot created successfully!", vbInformation

End Sub

' ============================================================
' Helper function for namespace filtering (supports wildcards)
' ============================================================
Function NamespaceMatches(ns As String, nsList As Variant) As Boolean
    Dim x As Variant
    For Each x In nsList
        If ns Like LCase(x) Then
            NamespaceMatches = True
            Exit Function
        End If
    Next x
    NamespaceMatches = False
End Function
