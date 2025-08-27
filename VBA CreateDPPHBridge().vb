Sub CreateDPPHBridge()
  Dim wsSource As Worksheet
  Dim selectedCell As Range
  Dim selectedDate As Variant
  
  ' Reference the source worksheet (first worksheet)
  Set wsSource = ThisWorkbook.Worksheets(1)
  
  ' Prompt user to select a cell containing the timestamp
  MsgBox "Please select the cell containing the timestamp you want to use.", vbInformation
  
  ' Allow user to select a cell
  On Error Resume Next
  Set selectedCell = Application.InputBox( _
    Prompt:="Click on the cell containing the timestamp (make sure the FIRST ROW has the timestamps):", _
    Title:="Select Timestamp Cell", _
    Type:=8)
  On Error GoTo 0
  
  ' Check if user canceled or selected an invalid cell
  If selectedCell Is Nothing Then
    MsgBox "Operation canceled by user.", vbInformation
    Exit Sub
  End If
  
  ' Get the value from the selected cell
  selectedDate = selectedCell.Value
  
  ' Check if the selected cell has a value
  If IsEmpty(selectedDate) Then
    MsgBox "The selected cell is empty. Please select a cell with a timestamp.", vbExclamation
    Exit Sub
  End If
  
  ' Call each table creation procedure with the selected date
  OverallTable selectedDate
  DSPTable selectedDate
  FlexTable selectedDate
  UTRTable selectedDate
  OTRTable selectedDate
  FTDSTable selectedDate
  OtherTable selectedDate
  DPPHBridge selectedDate
  
  MsgBox "DPPH Bridge for : " & selectedDate, vbInformation
End Sub

' B.Table Overall(DSP + Flex)
Sub OverallTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsDPPH As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "Overall (DSP + AmFlex)" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "Overall(DSP + AmFlex)" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'Overall(DSP + AmFlex)' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete Overall sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Overall").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new Overall sheet
    Set wsDPPH = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDPPH.Name = "Overall"
    
    ' Set up headers
    wsDPPH.Range("A1").Value = "Overall (DSP + AmFlex)"
    wsDPPH.Range("B1").Value = "Value for " & targetDate
    wsDPPH.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsDPPH.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics in the column under the bucket title
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsDPPH.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsDPPH.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsDPPH.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsDPPH.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsDPPH.Range("A1:B" & (rowCounter - 1)).Select
        wsDPPH.ListObjects.Add(xlSrcRange, wsDPPH.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "OverallTable"
    End If
    
    wsDPPH.Range("A1").Select
End Sub

' C.Table DSP
Sub DSPTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsDSP As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "DSP" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "DSP" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'DSP' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete DSP sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("DSP").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new DSP sheet
    Set wsDSP = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDSP.Name = "DSP"
    
    ' Set up headers
    wsDSP.Range("A1").Value = "DSP"
    wsDSP.Range("B1").Value = "Value for " & targetDate
    wsDSP.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsDSP.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsDSP.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsDSP.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsDSP.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsDSP.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsDSP.Range("A1:B" & (rowCounter - 1)).Select
        wsDSP.ListObjects.Add(xlSrcRange, wsDSP.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "DSPTable"
    End If
    
    wsDSP.Range("A1").Select
End Sub

' D.Table Flex
Sub FlexTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsFlex As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "AmFlex" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "AmFlex" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'AmFlex' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete AmFlex sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("AmFlex").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new AmFlex sheet
    Set wsFlex = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsFlex.Name = "AmFlex"
    
    ' Set up headers
    wsFlex.Range("A1").Value = "AmFlex"
    wsFlex.Range("B1").Value = "Value for " & targetDate
    wsFlex.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsFlex.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsFlex.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsFlex.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsFlex.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsFlex.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsFlex.Range("A1:B" & (rowCounter - 1)).Select
        wsFlex.ListObjects.Add(xlSrcRange, wsFlex.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "AmFlexTable"
    End If
    
    wsFlex.Range("A1").Select
End Sub

' B.Table UTR
Sub UTRTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsUTR As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "UTR" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "UTR" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'UTR' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete UTR sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("UTR").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new UTR sheet
    Set wsUTR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsUTR.Name = "UTR"
    
    ' Set up headers
    wsUTR.Range("A1").Value = "UTR"
    wsUTR.Range("B1").Value = "Value for " & targetDate
    wsUTR.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsUTR.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsUTR.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsUTR.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsUTR.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsUTR.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsUTR.Range("A1:B" & (rowCounter - 1)).Select
        wsUTR.ListObjects.Add(xlSrcRange, wsUTR.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "UTRTable"
    End If
    
    wsUTR.Range("A1").Select
End Sub

' E.Table OTR
Sub OTRTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsOTR As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "OTR" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "OTR" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'OTR' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete OTR sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("OTR").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new OTR sheet
    Set wsOTR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOTR.Name = "OTR"
    
    ' Set up headers
    wsOTR.Range("A1").Value = "OTR"
    wsOTR.Range("B1").Value = "Value for " & targetDate
    wsOTR.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsOTR.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsOTR.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsOTR.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsOTR.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsOTR.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsOTR.Range("A1:B" & (rowCounter - 1)).Select
        wsOTR.ListObjects.Add(xlSrcRange, wsOTR.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "OTRTable"
    End If
    
    wsOTR.Range("A1").Select
End Sub

' F.Table FTDS
Sub FTDSTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsFTDS As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "FTDS" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "FTDS" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'FTDS' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete FTDS sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("FTDS").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new FTDS sheet
    Set wsFTDS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsFTDS.Name = "FTDS"
    
    ' Set up headers
    wsFTDS.Range("A1").Value = "FTDS"
    wsFTDS.Range("B1").Value = "Value for " & targetDate
    wsFTDS.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsFTDS.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsFTDS.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsFTDS.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsFTDS.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsFTDS.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsFTDS.Range("A1:B" & (rowCounter - 1)).Select
        wsFTDS.ListObjects.Add(xlSrcRange, wsFTDS.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "FTDSTable"
    End If
    
    wsFTDS.Range("A1").Select
End Sub

' F.Table FTDS
Sub OtherTable(targetDate As Variant)
    Dim wsUserInput As Worksheet
    Dim wsTable As Worksheet
    Dim wsFTDS As Worksheet
    Dim dateColumn As Integer
    Dim bucketCol As Long
    
    ' Reference the required worksheets
    Set wsUserInput = ThisWorkbook.Worksheets("INPUT")
    Set wsTable = ThisWorkbook.Worksheets(1)
    
    ' Find the column with "Other" in row 4
    bucketCol = 0
    For c = 1 To wsUserInput.UsedRange.Columns.Count
        If wsUserInput.Cells(4, c).Value = "Other" Then
            bucketCol = c
            Exit For
        End If
    Next c
    
    If bucketCol = 0 Then
        MsgBox "Could not find 'Other' bucket in row 4!", vbExclamation
        Exit Sub
    End If
    
    ' Delete FTDS sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Other").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new FTDS sheet
    Set wsFTDS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsFTDS.Name = "Other"
    
    ' Set up headers
    wsFTDS.Range("A1").Value = "Other"
    wsFTDS.Range("B1").Value = "Value for " & targetDate
    wsFTDS.Range("A1:B1").Font.Bold = True
    
    ' Find the column in Table sheet that matches the target date
    dateColumn = 0
    For c = 1 To wsTable.UsedRange.Columns.Count
        If wsTable.Cells(1, c).Value = targetDate Then
            dateColumn = c
            Exit For
        End If
    Next c
    
    ' Check if date was found
    If dateColumn = 0 Then
        MsgBox "The selected date was not found in the first sheet!", vbExclamation
        wsFTDS.Range("A2").Value = "Date not found in first sheet"
        Exit Sub
    End If
    
    ' Loop through metrics and populate the sheet
    Dim rowCounter As Integer
    rowCounter = 2
    
    For r = 5 To wsUserInput.UsedRange.Rows.Count
        If IsEmpty(wsUserInput.Cells(r, bucketCol).Value) Then Exit For
        
        Dim metricName As String
        metricName = wsUserInput.Cells(r, bucketCol).Value
        
        ' Look for the metric in column B of Table sheet
        Dim foundCell As Range
        Set foundCell = wsTable.Columns(2).Find(What:=metricName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Add metric name and value to sheet
        wsFTDS.Cells(rowCounter, 1).Value = metricName
        
        If Not foundCell Is Nothing Then
            wsFTDS.Cells(rowCounter, 2).Value = wsTable.Cells(foundCell.Row, dateColumn).Value
        Else
            wsFTDS.Cells(rowCounter, 2).Value = "Metric not found"
        End If
        
        rowCounter = rowCounter + 1
    Next r
    
    ' Format the table
    wsFTDS.Columns("A:B").AutoFit
    If rowCounter > 2 Then
        wsFTDS.Range("A1:B" & (rowCounter - 1)).Select
        wsFTDS.ListObjects.Add(xlSrcRange, wsFTDS.Range("A1:B" & (rowCounter - 1)), , xlYes).Name = "FTDSTable"
    End If
    
    wsFTDS.Range("A1").Select
End Sub

Sub DPPHBridge(targetDate As Variant)
    Dim ws As Worksheet
    Dim bridgeSheetName As String
    Dim otherSheet As Worksheet
    Dim lastRow As Long, r As Long, currentRow As Long
    
    ' Create sheet name based on selected timestamp (limit to 31 characters)
    bridgeSheetName = Left(targetDate & " Bridge", 31)
    
    ' Delete sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(bridgeSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new sheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = bridgeSheetName
    
    ' Add content to the sheet
    With ws
        .Range("A1").Value = "DPPH Bridge for"
        .Range("B1").Value = targetDate
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").Font.Size = 14
        
        ' Use error handling for all references
        On Error Resume Next
        
        .Range("A2").Value = "Plan DPPH Sequenced Packages (AmFlex + DSP) Plan DPPH"
        .Range("B2").Value = Worksheets("Overall").Range("B3").Value
        .Range("A2:B2").Font.Bold = True
        
        .Range("A3").Value = "Actual DPPH (AmFlex + DSP)"
        .Range("B3").Value = Worksheets("Overall").Range("B4").Value
        .Range("A3:B3").Font.Bold = True
        
        .Range("A4").Value = "% PvA DPPH (DSP+AmFlex) [WWSC]"
        .Range("B4").Value = CDbl(Worksheets("Overall").Range("B2").Value) * 100
        .Range("A4:B4").Font.Bold = True
        
        .Range("A5").Value = "Sequenced Packages"
        .Range("B5").Value = Worksheets("Overall").Range("B13").Value
        
        .Range("A6").Value = "Delivered Packages"
        .Range("B6").Value = Worksheets("Overall").Range("B11").Value
        
        .Range("A7").Value = "RC1"
        .Range("A7").Font.Bold = True
        
        .Range("A8").Value = "Delivery Failure Leakage"
        .Range("B8").Value = Worksheets("Overall").Range("B10").Value
        
        .Range("A9").Value = "Delivery Failure Packages"
        .Range("B9").Value = Worksheets("Overall").Range("B14").Value
        
        .Range("A10").Value = "1.1. Package issues - Delivery Failure"
        .Range("B10").Value = Worksheets("OTR").Range("B9").Value
        
        .Range("A11").Value = "1.1.2 Delivery Attempted - Object Missing"
        .Range("B11").Value = Worksheets("OTR").Range("B15").Value
        
        .Range("A12").Value = "1.2. Customer issues - Delivery Failure"
        .Range("B12").Value = Worksheets("OTR").Range("B10").Value
        
        .Range("A13").Value = "1.2.1 Delivery Attempted - Customer Unavailable"
        .Range("B13").Value = Worksheets("OTR").Range("B18").Value
        
        .Range("A14").Value = "1.2.2 Delivery Attempted - UTA"
        .Range("B14").Value = Worksheets("OTR").Range("B16").Value
        
        .Range("A15").Value = "1.2.3 Delivery Attempted - UTL"
        .Range("B15").Value = Worksheets("OTR").Range("B17").Value
        
        .Range("A16").Value = "1.2.4 Fail Attempt - Rejected"
        .Range("B16").Value = Worksheets("OTR").Range("B24").Value
        
        .Range("A17").Value = "1.2.5 Delivery Attempted - Age Verification Fail"
        .Range("B17").Value = Worksheets("OTR").Range("B20").Value
        
        .Range("A18").Value = "1.3 Delivery Attempted - Business Closed"
        .Range("B18").Value = Worksheets("OTR").Range("B11").Value
        
        .Range("A19").Value = "1.4 Delivery Failure - Unknown Reason"
        .Range("B19").Value = Worksheets("OTR").Range("B13").Value
        
        .Range("A21").Value = "RC2"
        .Range("A21").Font.Bold = True
        
        .Range("A22").Value = "Dispatch Failure Leakage"
        .Range("B22").Value = Worksheets("Overall").Range("B9").Value
        
        .Range("A23").Value = "Dispatch Failure Packages"
        .Range("B23").Value = Worksheets("Overall").Range("B15").Value
        
        .Range("A24").Value = "2.1. Re-Dispatched more than Re-sequenced"
        .Range("B24").Value = Worksheets("UTR").Range("B8").Value
        
        .Range("A25").Value = "2.2. Sequenced but never Dispatched"
        .Range("B25").Value = Worksheets("UTR").Range("B7").Value
        
        .Range("A26").Value = "2.3. UTR package issues"
        .Range("B26").Value = Worksheets("UTR").Range("B2").Value
        
        .Range("A27").Value = "2.4. Not Attempted - Customer Reject"
        .Range("B27").Value = Worksheets("OTR").Range("B7").Value
        
        .Range("A28").Value = "2.5. Dispatch Failure - Unknown Reason"
        .Range("B28").Value = Worksheets("UTR").Range("B9").Value
        
        .Range("A29").Value = "2.6. Re-Sequenced more than Re-dispatched"
        .Range("B29").Value = Worksheets("UTR").Range("B11").Value
        
        .Range("A31").Value = "RC3"
        .Range("A31").Font.Bold = True
        
        .Range("A32").Value = "AmFlex DPPH"
        .Range("B32").Value = Worksheets("AmFlex").Range("B2").Value
        
        .Range("A33").Value = "Dispatched Shipment AmFlex"
        .Range("B33").Value = Worksheets("AmFlex").Range("B3").Value
        
        .Range("A34").Value = "Delivered Shipment AmFlex"
        .Range("B34").Value = Worksheets("AmFlex").Range("B4").Value
        
        .Range("A35").Value = "Delivery Failure Leakage AmFlex"
        .Range("B35").Value = Worksheets("AmFlex").Range("B9").Value
        
        .Range("A36").Value = "BWND Leakages AmFlex"
        .Range("B36").Value = Worksheets("AmFlex").Range("B7").Value
        
        ' Add Other section from Other sheet
        On Error Resume Next
        Set otherSheet = ThisWorkbook.Worksheets("Other")
        On Error GoTo 0
        
        If Not otherSheet Is Nothing Then
            lastRow = otherSheet.Cells(otherSheet.Rows.Count, "A").End(xlUp).Row
            
            ' Only add "Other" section if there are metrics (more than just the header row)
            If lastRow > 1 Then
                .Range("A38").Value = "Other"
                .Range("A38").Font.Bold = True
                
                currentRow = 39 ' Start adding metrics from row 39
                For r = 2 To lastRow ' Start from row 2 to skip header
                    .Range("A" & currentRow).Value = otherSheet.Range("A" & r).Value
                    .Range("B" & currentRow).Value = otherSheet.Range("B" & r).Value
                    currentRow = currentRow + 1
                Next r
            End If
        End If
        
        .Columns("A:B").AutoFit
    End With
    
    MsgBox "DPPH Bridge sheet has been created successfully for date: " & targetDate, vbInformation
End Sub


