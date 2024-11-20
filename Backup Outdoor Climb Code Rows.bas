Attribute VB_Name = "Module3"
Sub Individual_Lists_Based_By_Rows()
    Dim ws As Worksheet
    Dim dataLists As New Collection
    Dim colNum As Long
    Dim i, j, k, numCols, totalDataRows, currentRow, startRow, endStartRow, nextStartRow As Long
    Dim bool As Boolean
    
    ' set the worksheet
    Set ws = ThisWorkbook.Sheets("Send Data")
    
    numCols = 4 ' change # of column: grade, date, name, location
    
    colNum = 1 ' this is going to refer to column A then go to column E and I respectively

    ' Initialize variables
    totalDataRows = 0
    currentRow = 0
    endStartRow = 0
    nextStartRow = 2

    startRow = 32 ' other variables declared everytime

    For j = 1 To 3
        For k = 1 To 7
            
            Dim dataList() As Variant
        
            totalDataRows = 0 ' reset variables back to the starting position
            
            startRow = nextStartRow  ' Start from the first row of data
            
            endStartRow = startRow
            
            If ws.Cells(endStartRow, colNum + 1).value = "" Then
                endStartRow = endStartRow + 1 'makes a one value blank list otherwise
                ' make it so it still finds the next starting value for when V shows up if blank
             ' not blank start from the top
            End If
            
            
            bool = True
        
            Do While bool And ws.Cells(endStartRow, colNum + 1).value <> ""
                If InStr(ws.Cells(endStartRow + 1, colNum).value, "V") <> 0 Then
                    endStartRow = endStartRow + 1
                    nextStartRow = endStartRow
                    bool = False
                Else
                    endStartRow = endStartRow + 1 ' we always want the end start row to end one cell after the last data
                End If
                    ' don't increment endStartRow because we are
                    'looking forward bc the first data point to the next set is on the V-grade
            Loop
            ' endStartRow is inclusive (means where the lastdata point is for the V-grade)
        
            currentRow = endStartRow
            
            'Loop skips if all the data is filled and it hits the V-grade before a blank space in the prior
            Do While InStr(ws.Cells(currentRow, colNum).value, "V") = 0 And nextStartRow <> endStartRow
                currentRow = currentRow + 1
            Loop
            
            If currentRow <> endStartRow Then
                nextStartRow = currentRow
            End If
            ' in theory nextStartRow should equal currentRow + 1
            
            'for the list from startRow to endStartRow
            
            'then new list starting from nextStartRow
            
            
            
            ' Resize the dataList array
            totalDataRows = endStartRow - startRow
            currentRow = startRow
            
            ReDim dataList(1 To totalDataRows, 1 To numCols) ' numCols including 4 grade, date, name, location
            If totalDataRows >= 1 And ws.Cells(currentRow, colNum + 1).value <> "" Then
        
                ' Fill the dataList array
                For i = 1 To totalDataRows
                    dataList(i, 1) = ws.Cells(startRow, colNum).value
                    dataList(i, 2) = ws.Cells(currentRow, colNum + 1).value ' Dates from column B
                    dataList(i, 3) = ws.Cells(currentRow, colNum + 2).value ' Names from column C
                    dataList(i, 4) = ws.Cells(currentRow, colNum + 3).value ' Locations from column D
                    currentRow = currentRow + 1
                Next i
            Else
                'dataList(1, 1) = "" ' Dates from column B 'Can use if I need to make empty lists
                'dataList(1, 2) = "" ' Names from column C
                'dataList(1, 3) = "" ' Locations from column D
                'dataList(1, 4) = ""
            End If
            
        dataLists.Add dataList
        
        Next k
    
    colNum = colNum + 4 ' moving from Column A looking for V to Column E
    
    'reset variables to the top
    totalDataRows = 0
    currentRow = 0
    endStartRow = 0
    nextStartRow = 2
    startRow = 2 ' starting at 2 baby
    
    Next j
    
    Dim reorderedLists As Collection
    Set reorderedLists = New Collection
    
    ' Add lists to the reordered collection based on the desired sequence
    For i = 1 To 7
        reorderedLists.Add dataLists(i) ' 1
        reorderedLists.Add dataLists(i + 7) ' 8
        reorderedLists.Add dataLists(i + 14) ' 15
    Next i
    ' Now you can use reorderedLists as needed

    outputData = "Bouldering Sends Summary:" & vbCrLf
 
    ' Loop through each dataList in the collection
    For j = 1 To reorderedLists.Count
        dataList = reorderedLists(j) ' Get the dataList from the collection
        
        ' Prepare the output for this list
        If dataList(1, 1) <> "" Then
            
            outputData = "Bouldering sends @ " & dataList(1, 1) & ":" & vbCrLf
            For i = LBound(dataList, 1) To UBound(dataList, 1)
                outputData = outputData & dataList(i, 1) & " | " & dataList(i, 2) & " | " & dataList(i, 3) & " | " & dataList(i, 4) & vbCrLf
            Next i
        Else
            outputData = "No Bouldering Sends at this grade :("
        End If
        
        ' Show the output for the current list in a MsgBox
        MsgBox outputData, vbInformation, "Bouldering Data - List " & j
    Next j

End Sub

