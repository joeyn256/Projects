Attribute VB_Name = "Module3"
Sub Delete_Tables_Graph()
    Set output = ThisWorkbook.Worksheets("Sort")
    
    ' Loop through all PivotTables on the worksheet and delete them
    For Each pt In output.PivotTables
        pt.TableRange2.Clear ' This clears the entire Pivot Table
    Next pt
    
    ' Delete existing PivotCharts
    Dim pc As ChartObject
    For Each pc In output.ChartObjects
        pc.Delete
    Next pc
    
    output.Columns("H:Z").ColumnWidth = 10
    MsgBox "All Pivot Tables and/or Pivot Charts deleted."
End Sub
Function CreatePivotTable(table_Config As Integer) As Boolean
    
    startRow = 5 'redefine public
    Set output = ThisWorkbook.Sheets("Sort") 'redefine public
    
    Dim pivotTable As pivotTable
    Dim pivotRange As Range
    'Dim pivotField As pivotField
    Dim firstEmptyRow As Long
    Dim pivotDestination As Range
    Dim userChoice As Integer
    Dim newSheetName As String
    
    On Error GoTo ErrorHandler
    CreatePivotTable = False ' Default to False unless successful

    ' Clear existing PivotTables on the source worksheet
    Dim pt As pivotTable
    For Each pt In output.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    ' Delete existing PivotCharts
    Dim pc As ChartObject
    For Each pc In output.ChartObjects
        pc.Delete
    Next pc
    
    ' Determine the first empty row
    firstEmptyRow = output.Range("A" & startRow & ":A" & output.Rows.Count).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
    If firstEmptyRow = startRow Then
        MsgBox "No data to create a PivotTable.", vbExclamation
        Exit Function
    End If

    ' Set the data range for the PivotTable
    Set pivotRange = output.Range("A" & startRow - 1 & ":D" & firstEmptyRow - 1)
    Set pivotDestination = output.Range("H6") ' Default location for the PivotTable
    
    ' Ask user where to place the PivotTable
    userChoice = MsgBox("Do you want to create the PivotTable on a new worksheet?" & vbCrLf & _
                        "Click Yes for a new worksheet, or No to place it on the current sheet.", _
                        vbYesNoCancel + vbQuestion, "Choose Output Location")

    If userChoice = vbYes Then
        ' Create a new worksheet for the PivotTable
        newSheetName = InputBox("Enter the name for the new worksheet:", "New Worksheet Name")
        If newSheetName <> "" Then
            Dim newSheet As Worksheet
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newSheet.Name = newSheetName
            Set pivotDestination = newSheet.Range("A1")
            Set output = newSheet
        Else
            MsgBox "No name entered. Using current sheet instead.", vbInformation
        End If
    ElseIf userChoice = vbCancel Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Function
    End If

    ' Create the PivotTable
    Dim pivotCache As pivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange.Address(ReferenceStyle:=xlR1C1), Version:=8)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination.Address(ReferenceStyle:=xlR1C1), TableName:="Sends by Location and V-Grade", DefaultVersion:=8)

    If table_Config = 1 Then
        Call Table_Config_1(pivotTable)
    ElseIf table_Config = 2 Then
        Call Table_Config_2(pivotTable)
    Else
        MsgBox ("unknown table called")
    End If

    ' Optional formatting
    output.Columns("H:Z").ColumnWidth = 10

    CreatePivotTable = True ' Successfully completed
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Function
Function Table_Config_1(pivotTable As pivotTable)
    ' Configure PivotTable fields
    With pivotTable
        .PivotFields("Grade").Orientation = xlColumnField
        .PivotFields("Grade").Position = 1
        
        .PivotFields("Location").Orientation = xlRowField
        .PivotFields("Location").Position = 1
        
        With .PivotFields("Date")
            .Orientation = xlRowField
            .Position = 2
            .AutoGroup ' Group by Year and Month
        End With
        
        ' Hide unnecessary fields
        On Error Resume Next
        .PivotFields("Days (Date)").Orientation = xlHidden
        .PivotFields("Quarters (Date)").Orientation = xlHidden
        On Error GoTo 0
        
        ' Prioritize Years field if available
        On Error Resume Next
        .PivotFields("Years (Date)").Position = 1
        On Error GoTo 0
        
        ' Hide the original Date field
        .PivotFields("Date").Orientation = xlHidden

        ' Add data field
        With .PivotFields("Grade")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "0"
        End With
    End With
End Function
Function Table_Config_2(pivotTable As pivotTable)
    ' Configure PivotTable fields
    With pivotTable
        .PivotFields("Grade").Orientation = xlColumnField
        .PivotFields("Grade").Position = 1
        
         .PivotFields("Location").Orientation = xlRowField
         .PivotFields("Location").Position = 1
        
        ' Add data field
        With .PivotFields("Grade")
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "0"
        End With
    End With

End Function
Sub TestCreatePivotTable()
    If CreatePivotTable(1) Then
        MsgBox "PivotTable created successfully!", vbInformation
    End If
End Sub
Sub TestCreatePivotTable2()
    If CreatePivotTable(2) Then
        MsgBox "PivotTable created successfully!", vbInformation
    End If
End Sub
Sub TestCreatePivotChart()
    If CreatePivotChart(1, xlColumnClustered) Then
        '
    End If
End Sub
Sub TestCreatePivotChart2()
    If CreatePivotChart(2, xlColumnClustered) Then
        '
    End If
End Sub
Sub TestCreatePivotChart3()
    If CreatePivotChart(2, xlPie) Then
        '
    End If
End Sub
Sub TestCreatePivotChart4()
    If CreatePivotChart(2, xlArea) Then
        '
    End If
End Sub
Function CreatePivotChart(call_num As Integer, chart_type As String) As Boolean
    Dim chartDestination As Range
    Dim pivotChart As ChartObject
    Dim pivotTable As pivotTable
    
    Set output = ThisWorkbook.Sheets("Sort") 'redefine public
    
    CreatePivotChart = False ' Default to False unless successful
    
    ' Delete existing PivotCharts
    Dim pc As ChartObject
    For Each pc In output.ChartObjects
        pc.Delete
    Next pc
    
    If call_num = 1 Then
         If CreatePivotTable(1) Then
         '
         End If
    ElseIf call_num = 2 Then
         If CreatePivotTable(2) Then
         '
         End If
    Else
        MsgBox ("unknown table called")
    End If
    
    ' Create the PivotTable
    MsgBox "PivotTable created successfully!", vbInformation

    ' Get the last created PivotTable
    On Error Resume Next
    Set pivotTable = output.PivotTables("Sends by Location and V-Grade")
    On Error GoTo 0
    
    If pivotTable Is Nothing Then
        MsgBox "Failed to locate the created PivotTable.", vbCritical
        Exit Function
    End If
    
    If chart_type = xlPie Then
        pivotTable.PivotFields("Grade").Orientation = xlHidden
    End If
    
     If chart_type = xlArea Then
        pivotTable.PivotFields("Grade").Orientation = xlHidden
        pivotTable.PivotFields("Location").Orientation = xlHidden
        pivotTable.PivotFields("Grade").Orientation = xlRowField
        pivotTable.PivotFields("Grade").Position = 1
    End If
    
    ' Define the destination for the PivotChart if oncurrent worksheet or new
    If output.Name = "Sort" Then
        Set chartDestination = output.Range("E6") ' Adjust as needed for positioning
    Else
        Set chartDestination = output.Range("A10")
    End If
    
    ' Create the PivotChart
    Set pivotChart = output.ChartObjects.Add( _
        left:=chartDestination.left, _
        Width:=400, _
        Top:=chartDestination.Top, _
        Height:=300)
    
    With pivotChart.chart
        .SetSourceData Source:=pivotTable.TableRange2
        .ChartType = chart_type ' Example: Clustered Column chart
        .HasTitle = True
        .ChartTitle.Text = "Sends by Location and V-Grade"
    End With

    CreatePivotChart = True ' Default to False unless successful
    
     ' Optional formatting
    output.Columns("H:Z").ColumnWidth = 10
    
    MsgBox "PivotChart created successfully!", vbInformation
    
End Function
