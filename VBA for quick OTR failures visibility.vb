‘VBA for quick OTR failures visibility (by DSP)

Sub RUNALL()
    Dim selectedProvider As String
    Dim ws As Worksheet
    
    Set ws = Worksheets(1)
    selectedProvider = GetProviderSelectionOnce(ws)
    
    GroupByDSPTrackingIDs
    CreateVisualisation selectedProvider
    CreateScannableIDTables selectedProvider
End Sub

Function GetProviderSelectionOnce(ws As Worksheet) As String
    Dim lastRow As Long, providerCol As Integer
    Dim providers As Object, dataRange As Variant
    Dim i As Long, msg As String, userInput As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find provider column index
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If LCase(ws.Cells(1, i).Value) = "provider_company_short_code" Then
            providerCol = i
            Exit For
        End If
    Next i
    
    If providerCol = 0 Then
        MsgBox "Could not find 'provider_company_short_code' column!", vbExclamation
        Exit Function
    End If
    
    ' Create dictionary to store unique providers
    Set providers = CreateObject("Scripting.Dictionary")
    
    ' Read data in one go for efficiency
    dataRange = ws.Range(ws.Cells(2, providerCol), ws.Cells(lastRow, providerCol)).Value
    
    ' Add unique providers to dictionary
    For i = 1 To UBound(dataRange, 1)
        If Not IsEmpty(dataRange(i, 1)) Then
            providers(Trim(CStr(dataRange(i, 1)))) = 1
        End If
    Next i
    
    If providers.Count = 0 Then
        MsgBox "No providers found in the data!", vbExclamation
        Exit Function
    End If
    
    ' Create message for input box
    msg = "Available providers:" & vbCrLf & vbCrLf
    For Each key In providers.keys
        msg = msg & key & vbCrLf
    Next key
    
    msg = msg & vbCrLf & "Enter the provider name exactly as shown above, or leave blank for all providers:"
    
    ' Get user input
    userInput = InputBox(msg, "Select Provider")
    
    ' Validate user input
    If userInput = "" Then
        GetProviderSelectionOnce = ""
    ElseIf providers.Exists(userInput) Then
        GetProviderSelectionOnce = userInput
    Else
        MsgBox "Provider '" & userInput & "' not found. Please try again.", vbExclamation
        GetProviderSelectionOnce = GetProviderSelectionOnce(ws)
    End If
End Function

Sub GroupByDSPTrackingIDs()
    Dim ws As Worksheet, summaryWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim providerCol As Integer, scannableIDCol As Integer
    Dim providers As Object, counts As Object, scannableIDs As Object
    Dim dataRange As Variant, i As Long
    Dim currentProvider As String, currentScannableID As String
    
    Set providers = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set scannableIDs = CreateObject("Scripting.Dictionary")
    
    Set ws = Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find column indices
    For i = 1 To lastCol
        Select Case LCase(ws.Cells(1, i).Value)
            Case "provider_company_short_code": providerCol = i
            Case "scannable_id": scannableIDCol = i
        End Select
    Next i
    
    If providerCol = 0 Or scannableIDCol = 0 Then
        MsgBox "Could not find required columns!", vbExclamation
        Exit Sub
    End If
    
    ' Process data
    dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    
    For i = 1 To UBound(dataRange, 1)
        currentProvider = Trim(CStr(dataRange(i, providerCol)))
        currentScannableID = Trim(CStr(dataRange(i, scannableIDCol)))
        
        If currentProvider <> "" Then
            If Not providers.Exists(currentProvider) Then
                providers(currentProvider) = 1
                counts(currentProvider) = 1
                scannableIDs(currentProvider) = currentScannableID
            Else
                counts(currentProvider) = counts(currentProvider) + 1
                If InStr(scannableIDs(currentProvider), currentScannableID) = 0 Then
                    scannableIDs(currentProvider) = scannableIDs(currentProvider) & ", " & currentScannableID
                End If
            End If
        End If
    Next i
    
    ' Create summary worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("DSPs_Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set summaryWs = Worksheets.Add(After:=ws)
    summaryWs.Name = "DSPs_Summary"
    
    ' Add headers
    With summaryWs
        .Range("A1").Value = "DSP Summary for All Providers"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3:C3").Value = Array("Provider Company Short Code", "Count", "Scannable IDs")
        .Range("A3:C3").Font.Bold = True
        .Range("A3:C3").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Add data
    i = 4
    For Each key In providers.keys
        summaryWs.Cells(i, 1).Value = key
        summaryWs.Cells(i, 2).Value = counts(key)
        summaryWs.Cells(i, 3).Value = scannableIDs(key)
        i = i + 1
    Next key
    
    ' Format the results
    With summaryWs
        .Columns("A:C").AutoFit
        .Columns("C").ColumnWidth = 100
        .Range("A3:C" & (i - 1)).Borders.LineStyle = xlContinuous
        .Activate
    End With
    
    MsgBox providers.Count & " unique provider(s) found and summarized.", vbInformation
End Sub

Sub CreateVisualisation(Optional selectedProvider As String = "")
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim providerCol As Integer, cycleCol As Integer, addressTypeCol As Integer, eventDateTimeCol As Integer
    Dim dataRange As Variant, i As Long
    Dim currentProvider As String, currentCycle As String, currentAddressType As String, currentEventDateTime As String
    Dim providerCounts As Object, cycleCounts As Object, addressTypeCounts As Object, eventDateTimeCounts As Object
    Dim chartObj As ChartObject, chartTop As Double
    
    Set ws = Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find column indices
    For i = 1 To lastCol
        Select Case LCase(ws.Cells(1, i).Value)
            Case "provider_company_short_code": providerCol = i
            Case "cycle_name": cycleCol = i
            Case "address_type": addressTypeCol = i
            Case "event_datetime": eventDateTimeCol = i
        End Select
    Next i
    
    If providerCol = 0 Then
        MsgBox "Could not find 'provider_company_short_code' column!", vbExclamation
        Exit Sub
    End If
    
    If eventDateTimeCol = 0 Then
        MsgBox "Could not find 'event_datetime' column!", vbExclamation
        Exit Sub
    End If
    
    ' If no provider was passed, get provider selection from user
    If selectedProvider = "" Then
        selectedProvider = GetProviderSelectionOnce(ws)
    End If
    
    Set providerCounts = CreateObject("Scripting.Dictionary")
    Set cycleCounts = CreateObject("Scripting.Dictionary")
    Set addressTypeCounts = CreateObject("Scripting.Dictionary")
    Set eventDateTimeCounts = CreateObject("Scripting.Dictionary")
    
    ' Process data
    dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    
    For i = 1 To UBound(dataRange, 1)
        currentProvider = Trim(CStr(dataRange(i, providerCol)))
        
        If currentProvider <> "" And (selectedProvider = "" Or currentProvider = selectedProvider) Then
            ' Count providers
            If Not providerCounts.Exists(currentProvider) Then
                providerCounts(currentProvider) = 1
            Else
                providerCounts(currentProvider) = providerCounts(currentProvider) + 1
            End If
            
            ' Count cycles
            If cycleCol > 0 Then
                currentCycle = Trim(CStr(dataRange(i, cycleCol)))
                If currentCycle <> "" Then
                    If Not cycleCounts.Exists(currentCycle) Then
                        cycleCounts(currentCycle) = 1
                    Else
                        cycleCounts(currentCycle) = cycleCounts(currentCycle) + 1
                    End If
                End If
            End If
            
            ' Count address types
            If addressTypeCol > 0 Then
                currentAddressType = Trim(CStr(dataRange(i, addressTypeCol)))
                If currentAddressType <> "" Then
                    If Not addressTypeCounts.Exists(currentAddressType) Then
                        addressTypeCounts(currentAddressType) = 1
                    Else
                        addressTypeCounts(currentAddressType) = addressTypeCounts(currentAddressType) + 1
                    End If
                End If
            End If
            
            ' Count event datetimes
            currentEventDateTime = Trim(CStr(dataRange(i, eventDateTimeCol)))
            If currentEventDateTime <> "" Then
                If Not eventDateTimeCounts.Exists(currentEventDateTime) Then
                    eventDateTimeCounts(currentEventDateTime) = 1
                Else
                    eventDateTimeCounts(currentEventDateTime) = eventDateTimeCounts(currentEventDateTime) + 1
                End If
            End If
        End If
    Next i
    
    ' Create or clear visualization worksheet
    On Error Resume Next
    Set wsResult = Worksheets("Visualisation")
    On Error GoTo 0
    
    If wsResult Is Nothing Then
        Set wsResult = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsResult.Name = "Visualisation"
    Else
        wsResult.Cells.Clear
        For Each chartObj In wsResult.ChartObjects
            chartObj.Delete
        Next chartObj
    End If
    
    ' Set up title
    wsResult.Range("A1").Value = "Data Visualisation" & IIf(selectedProvider = "", " - All Providers", " - " & selectedProvider)
    wsResult.Range("A1").Font.Size = 14
    wsResult.Range("A1").Font.Bold = True
    
    ' Create provider count table if showing all providers
    If selectedProvider = "" Then
        wsResult.Range("A3").Value = "Provider Counts"
        wsResult.Range("A3").Font.Bold = True
        wsResult.Range("A4:B4").Value = Array("Provider", "Count")
        wsResult.Range("A4:B4").Font.Bold = True
        
        i = 5
        For Each key In providerCounts.keys
            wsResult.Cells(i, 1).Value = key
            wsResult.Cells(i, 2).Value = providerCounts(key)
            i = i + 1
        Next key
        
        ' Create provider pie chart
        Set chartObj = wsResult.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=300)
        With chartObj.Chart
            .SetSourceData Source:=wsResult.Range("A4:B" & (i - 1))
            .ChartType = xlPie
            .HasTitle = True
            .chartTitle.Text = "Distribution by Provider"
            .ApplyLayout 2
            .HasLegend = True
            .Legend.Position = xlLegendPositionRight
        End With
        
        chartTop = 350
    Else
        chartTop = 50
    End If
    
    ' Create cycle charts if data exists
    If cycleCol > 0 And cycleCounts.Count > 0 Then
        CreateCountChart wsResult, cycleCounts, "Cycle", "Distribution by Cycle", _
                        IIf(selectedProvider = "", "D", "A"), IIf(selectedProvider = "", 50, chartTop), _
                        IIf(selectedProvider = "", 750, 300)
    End If
    
    ' Create address type charts if data exists
    If addressTypeCol > 0 And addressTypeCounts.Count > 0 Then
        CreateCountChart wsResult, addressTypeCounts, "Address Type", "Distribution by Address Type", _
                        IIf(selectedProvider = "", "G", "D"), IIf(selectedProvider = "", chartTop, 50), _
                        IIf(selectedProvider = "", 300, 750)
    End If
    
    ' Create event datetime line chart
    CreateEventDateTimeLineChart wsResult, eventDateTimeCounts, "Event DateTime", "Events Over Time", _
                        IIf(selectedProvider = "", "J", "G"), IIf(selectedProvider = "", 50, chartTop + 350), _
                        IIf(selectedProvider = "", 750, 300)
    
    ' Format columns
    wsResult.Columns("A:L").ColumnWidth = 15
    wsResult.Activate
    
    MsgBox "Visualisation created successfully.", vbInformation
End Sub

Sub CreateCountChart(ws As Worksheet, countDict As Object, labelName As String, chartTitle As String, _
                    colLetter As String, topPosition As Double, leftPosition As Double)
    Dim i As Long, chartObj As ChartObject
    
    ws.Range(colLetter & "3").Value = labelName & " Counts"
    ws.Range(colLetter & "3").Font.Bold = True
    ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & "4").Value = Array(labelName, "Count")
    ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & "4").Font.Bold = True
    
    i = 5
    For Each key In countDict.keys
        ws.Range(colLetter & i).Value = key
        ws.Range(Chr(Asc(colLetter) + 1) & i).Value = countDict(key)
        i = i + 1
    Next key
    
    ' Create pie chart
    Set chartObj = ws.ChartObjects.Add(Left:=leftPosition, Top:=topPosition, Width:=400, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & (i - 1))
        .ChartType = xlPie
        .HasTitle = True
        .chartTitle.Text = chartTitle
        .ApplyLayout 2
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With
End Sub



Sub CreateEventDateTimeLineChart(ws As Worksheet, countDict As Object, labelName As String, chartTitle As String, _
 colLetter As String, topPosition As Double, leftPosition As Double)
 Dim i As Long, chartObj As ChartObject
 Dim sortedDates() As Variant
 Dim dateKey As Variant, idx As Long
 
 ws.Range(colLetter & "3").Value = labelName & " Counts"
 ws.Range(colLetter & "3").Font.Bold = True
 ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & "4").Value = Array(labelName, "Count")
 ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & "4").Font.Bold = True
 
 ' Sort the dates
 ReDim sortedDates(0 To countDict.Count - 1)
 idx = 0
 For Each dateKey In countDict.keys
 sortedDates(idx) = dateKey
 idx = idx + 1
 Next dateKey
 
 ' Simple bubble sort for dates - handle dd/mm/yyyy format
 Dim j As Long, temp As Variant
 For i = 0 To UBound(sortedDates) - 1
 For j = i + 1 To UBound(sortedDates)
 ' Convert string dates in dd/mm/yyyy format to proper date values
 If DateValue(sortedDates(i)) > DateValue(sortedDates(j)) Then
 temp = sortedDates(i)
 sortedDates(i) = sortedDates(j)
 sortedDates(j) = temp
 End If
 Next j
 Next i
 
 ' Write sorted data to worksheet
 i = 5
 For Each dateKey In sortedDates
 ' Format as proper Excel date - handle dd/mm/yyyy format
 ws.Range(colLetter & i).Value = DateValue(dateKey)
 ws.Range(colLetter & i).NumberFormat = "dd/mm/yyyy"
 ws.Range(Chr(Asc(colLetter) + 1) & i).Value = countDict(dateKey)
 i = i + 1
 Next dateKey
 
 ' Create line chart
 Set chartObj = ws.ChartObjects.Add(Left:=leftPosition, Top:=topPosition, Width:=600, Height:=300)
 With chartObj.Chart
 .SetSourceData Source:=ws.Range(colLetter & "4:" & Chr(Asc(colLetter) + 1) & (i - 1))
 .ChartType = xlLineMarkers ' Use line with markers for better visibility
 .HasTitle = True
 .chartTitle.Text = chartTitle
 
 ' Format the chart
 With .Axes(xlCategory)
 .HasTitle = True
 .AxisTitle.Text = "Date/Time"
 .CategoryType = xlTimeScale ' This is crucial for proper date handling
 
 ' Format the date labels in dd/mm/yyyy format
 .TickLabels.NumberFormat = "dd/mm/yyyy"
 
 ' Adjust label spacing if many data points
 If countDict.Count > 20 Then
 .TickLabelSpacing = countDict.Count \ 10
 End If
 End With
 
 With .Axes(xlValue)
 .HasTitle = True
 .AxisTitle.Text = "Count"
 .MinimumScale = 0 ' Start Y-axis at 0
 End With
 
 ' Add data labels if there aren't too many points
 If countDict.Count <= 20 Then
 .SeriesCollection(1).HasDataLabels = True
 .SeriesCollection(1).DataLabels.ShowValue = True
 End If
 
 ' Format the series
 With .SeriesCollection(1)
 .MarkerSize = 8
 .MarkerStyle = xlMarkerStyleCircle
 .Format.Line.Weight = 2
 End With
 End With
End Sub




Sub CreateScannableIDTables(Optional selectedProvider As String = "")
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim providerCol As Integer, scannableIDCol As Integer, cycleNameCol As Integer
    Dim addressTypeCol As Integer, dateCol As Integer
    Dim dateScanIDs As Object, cycleScanIDs As Object, addressTypeScanIDs As Object
    Dim i As Long, j As Long, k As Long
    
    ' Set references to worksheets
    Set ws = Worksheets(1)
    
    ' Create new worksheet for results
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Scannable IDs by DSP").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResult = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsResult.Name = "Scannable IDs by DSP"
    
    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find required columns
    For i = 1 To lastCol
        Select Case LCase(ws.Cells(1, i).Value)
            Case "provider_company_short_code": providerCol = i
            Case "scannable_id": scannableIDCol = i
            Case "cycle_name": cycleNameCol = i
            Case "address_type": addressTypeCol = i
            Case "date": dateCol = i
        End Select
    Next i
    
    ' Check if required columns were found
    If scannableIDCol = 0 Then
        MsgBox "Could not find 'scannable_id' column!", vbExclamation
        Exit Sub
    End If
    
    ' Create dictionaries
    Set dateScanIDs = CreateObject("Scripting.Dictionary")
    Set cycleScanIDs = CreateObject("Scripting.Dictionary")
    Set addressTypeScanIDs = CreateObject("Scripting.Dictionary")
    
    ' Process data
    For i = 2 To lastRow
        ' Skip if provider doesn't match (when filter is applied)
        If selectedProvider <> "" And ws.Cells(i, providerCol).Value <> selectedProvider Then
            GoTo NextRow
        End If
        
        ' Get scannable ID
        Dim scanID As String
        scanID = Trim(ws.Cells(i, scannableIDCol).Value)
        If scanID = "" Then GoTo NextRow
        
        ' Group by date if date column exists
        If dateCol > 0 Then
            Dim dateKey As String
            dateKey = Format(ws.Cells(i, dateCol).Value, "yyyy-mm-dd")
            If dateKey <> "" Then
                If Not dateScanIDs.Exists(dateKey) Then
                    dateScanIDs(dateKey) = scanID
                ElseIf InStr(dateScanIDs(dateKey), scanID) = 0 Then
                    dateScanIDs(dateKey) = dateScanIDs(dateKey) & ", " & scanID
                End If
            End If
        End If
        
        ' Group by cycle name if column exists
        If cycleNameCol > 0 Then
            Dim cycleKey As String
            cycleKey = Trim(ws.Cells(i, cycleNameCol).Value)
            If cycleKey <> "" Then
                If Not cycleScanIDs.Exists(cycleKey) Then
                    cycleScanIDs(cycleKey) = scanID
                ElseIf InStr(cycleScanIDs(cycleKey), scanID) = 0 Then
                    cycleScanIDs(cycleKey) = cycleScanIDs(cycleKey) & ", " & scanID
                End If
            End If
        End If
        
        ' Group by address type if column exists
        If addressTypeCol > 0 Then
            Dim addrKey As String
            addrKey = Trim(ws.Cells(i, addressTypeCol).Value)
            If addrKey <> "" Then
                If Not addressTypeScanIDs.Exists(addrKey) Then
                    addressTypeScanIDs(addrKey) = scanID
                ElseIf InStr(addressTypeScanIDs(addrKey), scanID) = 0 Then
                    addressTypeScanIDs(addrKey) = addressTypeScanIDs(addrKey) & ", " & scanID
                End If
            End If
        End If
NextRow:
    Next i
    
    ' Create tables in result worksheet
    wsResult.Cells(1, 1).Value = "Scannable IDs for " & IIf(selectedProvider = "", "All Providers", selectedProvider)
    wsResult.Cells(1, 1).Font.Bold = True
    wsResult.Cells(1, 1).Font.Size = 14
    
    ' Create table for scannable IDs by date
    i = 3
    If dateCol > 0 And dateScanIDs.Count > 0 Then
        wsResult.Cells(i, 1).Value = "By Date"
        wsResult.Cells(i, 1).Font.Bold = True
        wsResult.Cells(i + 1, 1).Value = "Date"
        wsResult.Cells(i + 1, 2).Value = "Scannable IDs"
        wsResult.Cells(i + 1, 3).Value = " "
        wsResult.Range("A" & (i + 1) & ":C" & (i + 1)).Font.Bold = True
        
        i = i + 2
        For Each key In dateScanIDs.keys
            wsResult.Cells(i, 1).Value = key
            wsResult.Cells(i, 2).Value = dateScanIDs(key)
            i = i + 1
        Next key
    End If
    
    ' Create table for scannable IDs by cycle name
    j = i + 1
    wsResult.Cells(j, 1).Value = "By Cycle Name"
    wsResult.Cells(j, 1).Font.Bold = True
    wsResult.Cells(j + 1, 1).Value = "Cycle Name"
    wsResult.Cells(j + 1, 2).Value = "Scannable IDs"
    wsResult.Cells(j + 1, 3).Value = " "
    wsResult.Range("A" & (j + 1) & ":C" & (j + 1)).Font.Bold = True
    
    j = j + 2
    For Each key In cycleScanIDs.keys
        wsResult.Cells(j, 1).Value = key
        wsResult.Cells(j, 2).Value = cycleScanIDs(key)
        j = j + 1
    Next key
    
    ' Create table for scannable IDs by address type if column exists
    If addressTypeCol > 0 And addressTypeScanIDs.Count > 0 Then
        k = j + 1
        wsResult.Cells(k, 1).Value = "By Address Type"
        wsResult.Cells(k, 1).Font.Bold = True
        wsResult.Cells(k + 1, 1).Value = "Address Type"
        wsResult.Cells(k + 1, 2).Value = "Scannable IDs"
        wsResult.Cells(k + 1, 3).Value = " "
        wsResult.Range("A" & (k + 1) & ":C" & (k + 1)).Font.Bold = True
        
        k = k + 2
        For Each key In addressTypeScanIDs.keys
            wsResult.Cells(k, 1).Value = key
            wsResult.Cells(k, 2).Value = addressTypeScanIDs(key)
            k = k + 1
        Next key
    End If
    
    ' Format columns
    wsResult.Columns("A:A").ColumnWidth = 20
    wsResult.Columns("B:B").ColumnWidth = 60
    wsResult.Columns("C:C").ColumnWidth = 20
    
    ' Format paste columns with light background color
    If dateScanIDs.Count > 0 Then
        wsResult.Range("C" & (3 + 2) & ":C" & (i - 1)).Interior.Color = RGB(240, 240, 240)
    End If
    wsResult.Range("C" & (j - cycleScanIDs.Count) & ":C" & (j - 1)).Interior.Color = RGB(240, 240, 240)
    If addressTypeCol > 0 And addressTypeScanIDs.Count > 0 Then
        wsResult.Range("C" & (k - addressTypeScanIDs.Count) & ":C" & (k - 1)).Interior.Color = RGB(240, 240, 240)
    End If
    
    MsgBox "Scannable ID tables for " & IIf(selectedProvider = "", "all providers", selectedProvider) & _
           " have been created in a new sheet named 'Scannable IDs by DSP'", vbInformation
End Sub


