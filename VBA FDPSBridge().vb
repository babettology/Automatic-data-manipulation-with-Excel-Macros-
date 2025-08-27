Sub FDPSBridge()
 Dim wsSource As Worksheet
 Dim wsSheet2 As Worksheet
 Dim selectedCell1 As Range
 Dim selectedCell2 As Range
 Dim selectedCell3 As Range
 Dim selectedDate1 As Variant
 Dim selectedDate2 As Variant
 Dim selectedDate3 As Variant
 Dim routedDateCol As Integer
 
 ' Reference the source worksheet (first worksheet) and Sheet 2
 Set wsSource = ThisWorkbook.Worksheets(1)
 Set wsSheet2 = ThisWorkbook.Worksheets(2)
 
 ' Allow user to select a cell for Details table
 On Error Resume Next
 Set selectedCell1 = Application.InputBox( _
 Prompt:="Click on the cell containing the timestamp for Details table (make sure the FIRST ROW has the timestamps):", _
 Title:="Select Timestamp Cell for Details", _
 Type:=8)
 On Error GoTo 0
 
 ' Check if user canceled or selected an invalid cell
 If selectedCell1 Is Nothing Then
 MsgBox "Operation canceled by user.", vbInformation
 Exit Sub
 End If
 
 ' Get the value from the selected cell
 selectedDate1 = selectedCell1.Value
 
 ' Check if the selected cell has a value
 If IsEmpty(selectedDate1) Then
 MsgBox "The selected cell is empty. Please select a cell with a timestamp.", vbExclamation
 Exit Sub
 End If
 
 ' Allow user to select a cell for WTD table
 On Error Resume Next
 Set selectedCell2 = Application.InputBox( _
 Prompt:="Click on the cell containing the timestamp for WTD table (make sure the FIRST ROW has the timestamps):", _
 Title:="Select Timestamp Cell for WTD", _
 Type:=8)
 On Error GoTo 0
 
 ' Check if user canceled or selected an invalid cell
 If selectedCell2 Is Nothing Then
 MsgBox "Operation canceled by user.", vbInformation
 Exit Sub
 End If
 
 ' Get the value from the selected cell
 selectedDate2 = selectedCell2.Value
 
 ' Check if the selected cell has a value
 If IsEmpty(selectedDate2) Then
 MsgBox "The selected cell is empty. Please select a cell with a timestamp.", vbExclamation
 Exit Sub
 End If
 
 ' Find the routed_date column in Sheet 2
 routedDateCol = 0
 For c = 1 To wsSheet2.UsedRange.Columns.count
   If wsSheet2.Cells(1, c).Value = "routed_date" Then
     routedDateCol = c
     Exit For
   End If
 Next c
 
 If routedDateCol = 0 Then
   MsgBox "Could not find 'routed_date' column in Sheet 2!", vbExclamation
   Exit Sub
 End If
 
 ' Create a collection of unique dates from the routed_date column
 Dim uniqueDates As New Collection
 Dim dateStr As String
 Dim i As Integer
 Dim alreadyExists As Boolean
 
 On Error Resume Next
 For i = 2 To wsSheet2.Cells(wsSheet2.Rows.count, routedDateCol).End(xlUp).Row
   dateStr = CStr(wsSheet2.Cells(i, routedDateCol).Value)
   If dateStr <> "" Then
     alreadyExists = False
     On Error Resume Next
     uniqueDates.Add dateStr, dateStr
     If Err.Number = 457 Then ' Key already exists
       alreadyExists = True
     End If
     On Error GoTo 0
   End If
 Next i
 
 ' Create a string with all unique dates for the InputBox
 Dim datesList As String
 datesList = "Available dates in routed_date column:" & vbCrLf & vbCrLf
 For i = 1 To uniqueDates.count
   datesList = datesList & i & ". " & uniqueDates(i) & vbCrLf
 Next i
 
 selectedDate3 = InputBox(datesList & vbCrLf & "Enter the date you want to use for Root Cause Analysis:", "Select Date for Root Causes")
 
 ' Check if user canceled or entered an invalid date
 If selectedDate3 = "" Then
   MsgBox "Operation canceled by user.", vbInformation
   Exit Sub
 End If
 
 ' Validate the entered date
 Dim validDate As Boolean
 validDate = False
 
 On Error Resume Next
 For i = 1 To uniqueDates.count
   If CStr(uniqueDates(i)) = selectedDate3 Then
     validDate = True
     Exit For
   End If
 Next i
 On Error GoTo 0
 
 If Not validDate Then
   MsgBox "The date you entered is not in the list of available dates.", vbExclamation
   Exit Sub
 End If
 
 ' Call each table creation procedure with the selected dates
 FormatColumn
 DetailsTable selectedDate1
 WTDTable selectedDate2
 RootCauseTable selectedDate3
 Bridge
 
 MsgBox "FDPS Bridge created successfully!" & vbCrLf & _
        "Details table: " & selectedDate1 & vbCrLf & _
        "WTD table: " & selectedDate2 & vbCrLf & _
        "Root Cause Analysis: " & selectedDate3, vbInformation
End Sub


' B.Table FTDS Details
Sub DetailsTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsFDPS As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "Details" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.count
        If wsUserInput.Cells(4, c).Value = "Details" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'Details' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete Overall sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Details").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new Overall sheet
    Set wsFDPS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsFDPS.Name = "Details"
    
    ' Set up headers
    wsFDPS.Range("A1").Value = "Details"
    wsFDPS.Range("B1").Value = "Values for " & targetDate
    wsFDPS.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsFDPS.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics in the column under the bucket title
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsFDPS.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsFDPS.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsFDPS.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsFDPS.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsFDPS.Range("A1:B" & (rowCounter - 1)).Select
        wsFDPS.ListObjects.Add(xlSrcRange, wsFDPS.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "DetailsTable"
    End If
    
   wsFDPS.Range("A1").Select
End Sub

Sub WTDTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsFDPS As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "WTD" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.count
        If wsUserInput.Cells(4, c).Value = "WTD" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'WTD' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    
    ' Delete WTD sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("WTD").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new WTD sheet
    Set wsFDPS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsFDPS.Name = "WTD"
    
    ' Set up headers
    wsFDPS.Range("A1").Value = "FDPS WTD"
    wsFDPS.Range("B1").Value = "Value for " & targetDate
    wsFDPS.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsFDPS.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics in the column under the bucket title
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsFDPS.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsFDPS.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsFDPS.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsFDPS.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsFDPS.Range("A1:B" & (rowCounter - 1)).Select
        wsFDPS.ListObjects.Add(xlSrcRange, wsFDPS.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "WTDTable"
    End If
    
   wsFDPS.Range("A1").Select
End Sub


Sub FormatColumn()
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim routedDateCol As Range
    Dim colIndex As Integer
    Dim cell As Range
    
    ' Set reference to Sheet2
    Set ws = ThisWorkbook.Sheets(2)
    
    ' Find the header row and locate the "routed_date" column
    Set headerRow = ws.Rows(1)
    colIndex = 0
    
    On Error Resume Next
    colIndex = Application.WorksheetFunction.Match("routed_date", headerRow, 0)
    On Error GoTo 0
    
    ' Check if column was found
    If colIndex = 0 Then
        MsgBox "Column 'routed_date' not found in Sheet2!", vbExclamation
        Exit Sub
    End If
    
    ' Get the range for the routed_date column (excluding header)
    Set routedDateCol = ws.Range(ws.Cells(2, colIndex), ws.Cells(ws.Rows.count, colIndex).End(xlUp))
    
    ' Loop through each cell and truncate to first 10 characters
    For Each cell In routedDateCol
        If Not IsEmpty(cell.Value) Then
            cell.Value = Left(cell.Value, 10)
        End If
    Next cell
    
    ' Also set the format to date format
    routedDateCol.NumberFormat = "dd/mm/yyyy"
    
    MsgBox "Values in 'routed_date' column have been truncated to 10 characters", vbInformation
End Sub

Sub RootCauseTable(targetDate As Variant)
    Dim wsUserInput As Worksheet, wsTable As Worksheet, wsRC As Worksheet
    Dim bucketCol As Long, metricName As String, metricCol As Integer
    Dim dateCol As Integer, primaryCol As Integer, secondaryCol As Integer
    Dim uniqueValues As Object, relationshipDict As Object
    Dim rowCounter As Integer, rowIndex As Long
    Dim r As Long, c As Long, totalCount As Long
    Dim key As Variant, cellValue As Variant
    Dim primaryReason As String, secondaryReason As String
    Dim trackingCol As Integer, providerCol As Integer, transporterCol As Integer
    Dim businessHoursCol As Integer, plannedArrivalCol As Integer, distanceCol As Integer
    Dim routeCodeCol As Integer  ' Added dedicated variable for route_code
    Dim detailsAdded As Boolean
    
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(2)
    
    For c = 1 To wsUserInput.UsedRange.Columns.count
        If wsUserInput.Cells(4, c).Value = "RC" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'RC' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    For c = 1 To wsTable.UsedRange.Columns.count
        If wsTable.Cells(1, c).Value = "routed_date" Then
            dateCol = c
        ElseIf wsTable.Cells(1, c).Value = "pickup_failure_reason" Then
            primaryCol = c
        ElseIf wsTable.Cells(1, c).Value = "secondary_pickup_failure_reason" Then
            secondaryCol = c
        ElseIf wsTable.Cells(1, c).Value = "tracking_id" Then
            trackingCol = c
        ElseIf wsTable.Cells(1, c).Value = "route_code" Then
            routeCodeCol = c  ' Corrected assignment to proper variable
        ElseIf wsTable.Cells(1, c).Value = "provider_company_short_code" Then
            providerCol = c
        ElseIf wsTable.Cells(1, c).Value = "transporter_id" Then
            transporterCol = c
        ElseIf wsTable.Cells(1, c).Value = "business_hours" Then
            businessHoursCol = c
        ElseIf wsTable.Cells(1, c).Value = "planned_arrival" Then
            plannedArrivalCol = c
        ElseIf wsTable.Cells(1, c).Value = "distance_in_meters" Then
            distanceCol = c
        End If
    Next c
    
    If dateCol = 0 Then
        MsgBox "Could not find 'routed_date' column in Sheet 2!", vbExclamation
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("RC").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsRC = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsRC.Name = "RC"
    
    wsRC.Range("A1").Value = "FDPS Root Cause Analysis"
    wsRC.Range("A1").Font.Bold = True
    wsRC.Range("A1").Font.Size = 14
    wsRC.Range("B1").Value = "Value for " & targetDate
    wsRC.Range("B1").Font.Bold = True
    wsRC.Range("B1").Font.Size = 14
    
    rowCounter = 3
    
    ' First process pickup_failure_reason
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    For rowIndex = 2 To wsTable.UsedRange.Rows.count
        If Left(wsTable.Cells(rowIndex, dateCol).Value, 10) = Left(targetDate, 10) Then
            cellValue = wsTable.Cells(rowIndex, primaryCol).Value
            If Not IsEmpty(cellValue) Then
                If uniqueValues.Exists(cellValue) Then
                    uniqueValues(cellValue) = uniqueValues(cellValue) + 1
                Else
                    uniqueValues.Add cellValue, 1
                End If
            End If
        End If
    Next rowIndex
    
    wsRC.Cells(rowCounter, 1).Value = "pickup_failure_reason"
    wsRC.Cells(rowCounter, 1).Font.Bold = True
    rowCounter = rowCounter + 1
    
    wsRC.Cells(rowCounter, 1).Value = "Value"
    wsRC.Cells(rowCounter, 2).Value = "Count"
    wsRC.Cells(rowCounter, 3).Value = "Percentage"
    wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
    rowCounter = rowCounter + 1
    
    totalCount = 0
    If uniqueValues.count > 0 Then
        totalCount = Application.WorksheetFunction.Sum(uniqueValues.Items)
    End If
    
    If uniqueValues.count > 0 Then
        For Each key In uniqueValues.Keys
            wsRC.Cells(rowCounter, 1).Value = key
            wsRC.Cells(rowCounter, 2).Value = uniqueValues(key)
            wsRC.Cells(rowCounter, 3).Value = uniqueValues(key) / totalCount
            wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
            rowCounter = rowCounter + 1
        Next key
        
        wsRC.Cells(rowCounter, 1).Value = "Total"
        wsRC.Cells(rowCounter, 2).Value = totalCount
        wsRC.Cells(rowCounter, 3).Value = 1
        wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
        wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
    Else
        wsRC.Cells(rowCounter, 1).Value = "No data for selected date"
        wsRC.Cells(rowCounter, 1).Font.Italic = True
    End If
    
    rowCounter = rowCounter + 3
    
    ' Then process secondary_pickup_failure_reason
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    For rowIndex = 2 To wsTable.UsedRange.Rows.count
        If Left(wsTable.Cells(rowIndex, dateCol).Value, 10) = Left(targetDate, 10) Then
            cellValue = wsTable.Cells(rowIndex, secondaryCol).Value
            If Not IsEmpty(cellValue) Then
                If uniqueValues.Exists(cellValue) Then
                    uniqueValues(cellValue) = uniqueValues(cellValue) + 1
                Else
                    uniqueValues.Add cellValue, 1
                End If
            End If
        End If
    Next rowIndex
    
    wsRC.Cells(rowCounter, 1).Value = "secondary_pickup_failure_reason"
    wsRC.Cells(rowCounter, 1).Font.Bold = True
    rowCounter = rowCounter + 1
    
    wsRC.Cells(rowCounter, 1).Value = "Value"
    wsRC.Cells(rowCounter, 2).Value = "Count"
    wsRC.Cells(rowCounter, 3).Value = "Percentage"
    wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
    rowCounter = rowCounter + 1
    
    totalCount = 0
    If uniqueValues.count > 0 Then
        totalCount = Application.WorksheetFunction.Sum(uniqueValues.Items)
    End If
    
    If uniqueValues.count > 0 Then
        For Each key In uniqueValues.Keys
            wsRC.Cells(rowCounter, 1).Value = key
            wsRC.Cells(rowCounter, 2).Value = uniqueValues(key)
            wsRC.Cells(rowCounter, 3).Value = uniqueValues(key) / totalCount
            wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
            rowCounter = rowCounter + 1
        Next key
        
        wsRC.Cells(rowCounter, 1).Value = "Total"
        wsRC.Cells(rowCounter, 2).Value = totalCount
        wsRC.Cells(rowCounter, 3).Value = 1
        wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
        wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
    Else
        wsRC.Cells(rowCounter, 1).Value = "No data for selected date"
        wsRC.Cells(rowCounter, 1).Font.Italic = True
    End If
    
    rowCounter = rowCounter + 3
    
    ' Then process Primary | Secondary Reason relationship
    Set relationshipDict = CreateObject("Scripting.Dictionary")
    For rowIndex = 2 To wsTable.UsedRange.Rows.count
        If Left(wsTable.Cells(rowIndex, dateCol).Value, 10) = Left(targetDate, 10) Then
            primaryReason = wsTable.Cells(rowIndex, primaryCol).Value
            secondaryReason = wsTable.Cells(rowIndex, secondaryCol).Value
            
            If Not IsEmpty(primaryReason) And Not IsEmpty(secondaryReason) Then
                key = primaryReason & " | " & secondaryReason
                
                If relationshipDict.Exists(key) Then
                    relationshipDict(key) = relationshipDict(key) + 1
                Else
                    relationshipDict.Add key, 1
                End If
            End If
        End If
    Next rowIndex
    
    wsRC.Cells(rowCounter, 1).Value = "Primary | Secondary Reason"
    wsRC.Cells(rowCounter, 1).Font.Bold = True
    rowCounter = rowCounter + 1
    
    wsRC.Cells(rowCounter, 1).Value = "Value"
    wsRC.Cells(rowCounter, 2).Value = "Count"
    wsRC.Cells(rowCounter, 3).Value = "Percentage"
    wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
    rowCounter = rowCounter + 1
    
    totalCount = 0
    If relationshipDict.count > 0 Then
        totalCount = Application.WorksheetFunction.Sum(relationshipDict.Items)
    End If
    
    If relationshipDict.count > 0 Then
        For Each key In relationshipDict.Keys
            wsRC.Cells(rowCounter, 1).Value = key
            wsRC.Cells(rowCounter, 2).Value = relationshipDict(key)
            wsRC.Cells(rowCounter, 3).Value = relationshipDict(key) / totalCount
            wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
            rowCounter = rowCounter + 1
        Next key
        
        wsRC.Cells(rowCounter, 1).Value = "Total"
        wsRC.Cells(rowCounter, 2).Value = totalCount
        wsRC.Cells(rowCounter, 3).Value = 1
        wsRC.Cells(rowCounter, 3).NumberFormat = "0.0%"
        wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 3)).Font.Bold = True
        
        ' Add the details table for Primary | Secondary Reason combinations
        rowCounter = rowCounter + 3
        
        wsRC.Cells(rowCounter, 1).Value = "Primary | Secondary Reason Details"
        wsRC.Cells(rowCounter, 1).Font.Bold = True
        rowCounter = rowCounter + 1
        
        wsRC.Cells(rowCounter, 1).Value = "Primary | Secondary"
        wsRC.Cells(rowCounter, 2).Value = "tracking_id"
        wsRC.Cells(rowCounter, 3).Value = "provider_company_short_code"
        wsRC.Cells(rowCounter, 4).Value = "transporter_id"
        wsRC.Cells(rowCounter, 5).Value = "business_hours"
        wsRC.Cells(rowCounter, 6).Value = "planned_arrival"
        wsRC.Cells(rowCounter, 7).Value = "distance_in_meters"
        wsRC.Cells(rowCounter, 8).Value = "route_code"
        wsRC.Range(wsRC.Cells(rowCounter, 1), wsRC.Cells(rowCounter, 8)).Font.Bold = True
        rowCounter = rowCounter + 1
        
        For Each key In relationshipDict.Keys
            detailsAdded = False
            For rowIndex = 2 To wsTable.UsedRange.Rows.count
                If Left(wsTable.Cells(rowIndex, dateCol).Value, 10) = Left(targetDate, 10) Then
                    primaryReason = wsTable.Cells(rowIndex, primaryCol).Value
                    secondaryReason = wsTable.Cells(rowIndex, secondaryCol).Value
                    
                    If Not IsEmpty(primaryReason) And Not IsEmpty(secondaryReason) Then
                        If primaryReason & " | " & secondaryReason = key Then
                            wsRC.Cells(rowCounter, 1).Value = key
                            
                            If trackingCol > 0 Then
                                wsRC.Cells(rowCounter, 2).Value = wsTable.Cells(rowIndex, trackingCol).Value
                            End If
                            
                            If providerCol > 0 Then
                                wsRC.Cells(rowCounter, 3).Value = wsTable.Cells(rowIndex, providerCol).Value
                            End If
                            
                            If transporterCol > 0 Then
                                wsRC.Cells(rowCounter, 4).Value = wsTable.Cells(rowIndex, transporterCol).Value
                            End If
                            
                            If businessHoursCol > 0 Then
                                wsRC.Cells(rowCounter, 5).Value = wsTable.Cells(rowIndex, businessHoursCol).Value
                            End If
                            
                            If plannedArrivalCol > 0 Then
                                wsRC.Cells(rowCounter, 6).Value = wsTable.Cells(rowIndex, plannedArrivalCol).Value
                            End If
                            
                            If distanceCol > 0 Then
    ' Get the distance value
    Dim distanceValue As Variant
    distanceValue = wsTable.Cells(rowIndex, distanceCol).Value
    
    ' Check if it's a number before rounding
    If IsNumeric(distanceValue) Then
        ' Round to 2 decimal places
        wsRC.Cells(rowCounter, 7).Value = Round(distanceValue, 2)
    Else
        ' If not a number, just use the original value
        wsRC.Cells(rowCounter, 7).Value = distanceValue
    End If
End If
                            
                            If routeCodeCol > 0 Then  ' Added route_code to the output
                                wsRC.Cells(rowCounter, 8).Value = wsTable.Cells(rowIndex, routeCodeCol).Value
                            End If
                            
                            rowCounter = rowCounter + 1
                            detailsAdded = True
                        End If
                    End If
                End If
            Next rowIndex
            
            If Not detailsAdded Then
                wsRC.Cells(rowCounter, 1).Value = key
                wsRC.Cells(rowCounter, 2).Value = "No details available"
                rowCounter = rowCounter + 1
            End If
            
            rowCounter = rowCounter + 1
        Next key
    Else
        wsRC.Cells(rowCounter, 1).Value = "No data for selected date"
        wsRC.Cells(rowCounter, 1).Font.Italic = True
    End If
    
    ' The incomplete For loop at the end was removed
End Sub


Sub Bridge()
    Dim wsUserInput As Worksheet, wsTable As Worksheet, wsRC As Worksheet
    Dim bucketCol As Long, metricName As String, metricCol As Integer
    Dim dateCol As Integer, primaryCol As Integer, secondaryCol As Integer
    Dim uniqueValues As Object, relationshipDict As Object
    Dim rowCounter As Integer, rowIndex As Long
    Dim r As Long, c As Long, totalCount As Long
    Dim key As Variant, cellValue As Variant
    Dim wsFDPSBridge As Worksheet
    Dim detailsSheet As Worksheet
    Dim fdpsValue As String
    Dim sheetExists As Boolean
    Dim rcSheet As Worksheet
    Dim startRow As Long, endRow As Long
    Dim i As Long, outputRow As Long
    
    ' Column indices for detailed information
    Dim trackingCol As Integer, providerCol As Integer, transporterCol As Integer
    Dim businessHoursCol As Integer, plannedArrivalCol As Integer, distanceCol As Integer
    Dim routeCodeCol As Integer
    
    ' Check if "FDPS Bridge" sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "FDPS Bridge" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        Sheets("FDPS Bridge").Delete
        Application.DisplayAlerts = True
    End If
    
    Set wsFDPSBridge = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsFDPSBridge.Name = "FDPS Bridge"
    
    ' Add content to the sheet
    With wsFDPSBridge
        .Range("A1").Value = "FDPS Bridge"
        .Range("B1").Value = ThisWorkbook.Sheets("Details").Range("B1").Value
        
        .Range("A2").Value = "DXM2 FDPS Compliance"
        .Range("B2").Value = Format(ThisWorkbook.Sheets("Details").Range("B2").Value * 100, "0.00") & "%"
        
        .Range("A3").Value = "FDPS (WTD)"
        .Range("B3").Value = Format(ThisWorkbook.Sheets("WTD").Range("B2").Value * 100, "0.00") & "%"
        
        .Range("A4").Value = "Total shipments impacted"
        .Range("B4").Value = ThisWorkbook.Sheets("Details").Range("B5").Value
        
        .Range("A5").Value = "Total shipments picked up"
        .Range("B5").Value = ThisWorkbook.Sheets("Details").Range("B6").Value
        
        .Range("A6").Value = "Total failures"
        .Range("B6").Value = ThisWorkbook.Sheets("Details").Range("B3").Value
        
        .Range("A7").Value = "FDPS Package Exclusions"
        .Range("B7").Value = ThisWorkbook.Sheets("Details").Range("B4").Value
        
        ' Format the header section
        .Range("A1:B3").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2:A3").Font.Size = 12
        
        ' Check if RC sheet exists
        On Error Resume Next
        Set rcSheet = ThisWorkbook.Sheets("RC")
        On Error GoTo 0
        
        If Not rcSheet Is Nothing Then
            ' Find the "Primary | Secondary Reason" section in RC sheet
            startRow = 0
            For i = 1 To rcSheet.UsedRange.Rows.count
                If rcSheet.Cells(i, 1).Value = "Primary | Secondary Reason" Then
                    startRow = i + 2 ' Skip the header row
                    Exit For
                End If
            Next i
            
            If startRow > 0 Then
                ' Find the end of this section
                endRow = startRow
                While Not IsEmpty(rcSheet.Cells(endRow, 1).Value) And rcSheet.Cells(endRow, 1).Value <> "Total"
                    endRow = endRow + 1
                Wend
                
                ' Add header for the table
                .Range("A9").Value = "Root Cause Analysis"
                .Range("A9").Font.Bold = True
                .Range("A9").Font.Size = 12
                
                ' Find the "Primary | Secondary Reason Details" section in RC sheet
                Dim detailsStartRow As Long
                detailsStartRow = 0
                For i = 1 To rcSheet.UsedRange.Rows.count
                    If rcSheet.Cells(i, 1).Value = "Primary | Secondary Reason Details" Then
                        detailsStartRow = i + 2 ' Skip the header row
                        Exit For
                    End If
                Next i
                
                If detailsStartRow > 0 Then
                    ' Process each Primary | Secondary reason
                    outputRow = 11
                    For i = startRow To endRow - 1
                        Dim currentReason As String, reasonCount As Long
                        currentReason = rcSheet.Cells(i, 1).Value
                        reasonCount = rcSheet.Cells(i, 2).Value
                        
                        ' Add header for this root cause
                        .Cells(outputRow, 1).Value = "Root Cause: " & currentReason & " – (" & reasonCount & " events)"
                        .Cells(outputRow, 1).Font.Bold = True
                        .Range(.Cells(outputRow, 1), .Cells(outputRow, 4)).Merge
                        outputRow = outputRow + 1
                        
                        ' Find all details for this reason in the details section
                        Dim detailsFound As Integer
                        detailsFound = 0
                        
                        ' Search through the details section for matching entries
                        Dim j As Long
                        For j = detailsStartRow To rcSheet.UsedRange.Rows.count
                            If rcSheet.Cells(j, 1).Value = currentReason Then
                                ' Format the detailed information
                                Dim detailText As String
                                detailText = "Details of instance " & (detailsFound + 1) & ":" & vbCrLf & _
                                            rcSheet.Cells(j, 2).Value & " (" & _
                                            rcSheet.Cells(j, 3).Value & " – " & _
                                            rcSheet.Cells(j, 8).Value & ", " & _
                                            rcSheet.Cells(j, 4).Value & ") : planned at " & _
                                            rcSheet.Cells(j, 6).Value & " (with time windows " & _
                                            rcSheet.Cells(j, 5).Value & "). Marked " & _
                                            rcSheet.Cells(j, 7).Value & "m from location."
                                
                                ' Add the detailed text to cell
                                .Cells(outputRow, 1).Value = detailText
                                .Cells(outputRow, 1).WrapText = True
                                .Cells(outputRow, 1).RowHeight = 60 ' Adjust row height
                                .Range(.Cells(outputRow, 1), .Cells(outputRow, 4)).Merge ' Merge cells for better display
                                .Cells(outputRow, 1).Font.Size = 10
                                
                                ' Add some formatting to highlight key information
                                .Cells(outputRow, 1).Characters(1, Len("Details of instance " & (detailsFound + 1) & ":")).Font.Bold = True
                                .Cells(outputRow, 1).Characters(1, Len("Details of instance " & (detailsFound + 1) & ":")).Font.Italic = True
                                
                                ' Add alternating colors for better readability
                                If detailsFound Mod 2 = 0 Then
                                    .Range(.Cells(outputRow, 1), .Cells(outputRow, 4)).Interior.Color = RGB(245, 245, 245)
                                End If
                                
                                outputRow = outputRow + 1
                                detailsFound = detailsFound + 1
                            ElseIf detailsFound > 0 And IsEmpty(rcSheet.Cells(j, 1).Value) Then
                                ' We've found all details for this reason
                                Exit For
                            End If
                        Next j
                        
                        If detailsFound = 0 Then
                            .Cells(outputRow, 1).Value = "No detailed information found for this root cause"
                            .Cells(outputRow, 1).Font.Italic = True
                            outputRow = outputRow + 1
                        End If
                        
                        ' Add a separator after each root cause
                        .Cells(outputRow, 1).Value = String(50, "-")
                        .Range(.Cells(outputRow, 1), .Cells(outputRow, 4)).Merge
                        .Cells(outputRow, 1).HorizontalAlignment = xlCenter
                        outputRow = outputRow + 2
                    Next i
                    
                    ' Add summary table header
                    .Cells(outputRow, 1).Value = "Root Cause Summary"
                    .Cells(outputRow, 1).Font.Bold = True
                    .Cells(outputRow, 1).Font.Size = 12
                    outputRow = outputRow + 2
                    
                    .Cells(outputRow, 1).Value = "Primary | Secondary Reason"
                    .Cells(outputRow, 2).Value = "Count"
                    .Cells(outputRow, 3).Value = "Percentage"
                    .Range(.Cells(outputRow, 1), .Cells(outputRow, 3)).Font.Bold = True
                    .Range(.Cells(outputRow, 1), .Cells(outputRow, 3)).Interior.Color = RGB(220, 230, 241) ' Light blue header
                    outputRow = outputRow + 1
                    
                    ' Copy the summary data
                    For i = startRow To endRow
                        .Cells(outputRow, 1).Value = rcSheet.Cells(i, 1).Value
                        .Cells(outputRow, 2).Value = rcSheet.Cells(i, 2).Value
                        .Cells(outputRow, 3).Value = rcSheet.Cells(i, 3).Value
                        .Cells(outputRow, 3).NumberFormat = "0.0%"
                        
                        ' Bold the total row
                        If rcSheet.Cells(i, 1).Value = "Total" Then
                            .Range(.Cells(outputRow, 1), .Cells(outputRow, 3)).Font.Bold = True
                        End If
                        
                        outputRow = outputRow + 1
                    Next i
                    
                    ' Format the table
                    .Range("A" & (outputRow - (endRow - startRow + 1)) & ":C" & (outputRow - 1)).Borders.Weight = xlThin
                    .Columns("A:C").AutoFit
                    
                    ' Add alternating row colors for better readability
                    For i = (outputRow - (endRow - startRow)) To (outputRow - 2) Step 2
                        .Range(.Cells(i, 1), .Cells(i, 3)).Interior.Color = RGB(240, 240, 240) ' Light gray for alternating rows
                    Next i
                Else
                    .Range("A11").Value = "Primary | Secondary Reason details not found in RC sheet"
                    .Range("A11").Font.Italic = True
                End If
            Else
                .Range("A11").Value = "Primary | Secondary Reason section not found in RC sheet"
                .Range("A11").Font.Italic = True
            End If
        Else
            .Range("A11").Value = "RC sheet not found"
            .Range("A11").Font.Italic = True
        End If
    End With
End Sub

