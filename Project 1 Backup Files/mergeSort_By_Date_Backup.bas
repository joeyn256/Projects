Attribute VB_Name = "Module4"
Sub mergeSort_By_Date_Backup()

Dim dataLists As New Collection

Set dataLists = list_By_Column() ' use the sort by column, because it is one less function and we are sorting by date anyways

'Recursive merge on reorderedLists
    Dim sortedList As Variant
    sortedList = RecursiveMerge(dataLists, 1, dataLists.Count)
    
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sort")
    
    Dim l As Long
    Dim t As Long
    t = 6
    ws1.Cells.Clear ' Clear any previous data on Sheet2
    'ws1.Rows("6:" & ws.Rows.Count).ClearContents ' Clear only rows below row 5
    
    ' Loop through sortedList and write values to Sheet2
    For l = LBound(sortedList, 1) To UBound(sortedList, 1)
        If sortedList(l, 1) <> "" Then
            ws1.Cells(t, 1).value = sortedList(l, 1) ' Column A
            ws1.Cells(t, 2).value = sortedList(l, 2) ' Column B
            ws1.Cells(t, 3).value = sortedList(l, 3) ' Column C
            ws1.Cells(t, 4).value = sortedList(l, 4) ' Column D
            t = t + 1
        Else
        
        End If
        
    Next l
    
    MsgBox "Successfully Sorted Climbs By Date", vbInformation
    
    ' Output sorted dates in the immediate debugger window
    Debug.Print "Final Sorted List:"
    Dim c As Integer
    For c = LBound(sortedList, 1) To UBound(sortedList, 1)
        Debug.Print sortedList(c, 2) ' Print the date in sorted order
    Next c
End Sub

Function RecursiveMerge(lists As Collection, left As Long, right As Long) As Variant
    If left = right Then
        Debug.Print "Single List at position "; left
        PrintArray lists(left)
        RecursiveMerge = lists(left)
        Exit Function
    End If
    
    Dim mid As Long
    mid = (left + right) \ 2
    
    Dim leftMerged As Variant
    Dim rightMerged As Variant
    leftMerged = RecursiveMerge(lists, left, mid)
    rightMerged = RecursiveMerge(lists, mid + 1, right)
    
    RecursiveMerge = Merge(leftMerged, rightMerged)
End Function

Function Merge(leftList As Variant, rightList As Variant) As Variant
    Dim i As Long, j As Long, k As Long
    Dim leftCount As Long, rightCount As Long
    leftCount = UBound(leftList, 1)
    rightCount = UBound(rightList, 1)
    
    Dim mergedList() As Variant
    ReDim mergedList(1 To leftCount + rightCount, 1 To 4)
    
    i = 1: j = 1: k = 1
    
    Debug.Print "Merging two lists:"
    Debug.Print "Left List:"
    PrintArray leftList
    Debug.Print "Right List:"
    PrintArray rightList
    
    Do While i <= leftCount And j <= rightCount
        If leftList(i, 2) <= rightList(j, 2) Then
            mergedList(k, 1) = leftList(i, 1)
            mergedList(k, 2) = leftList(i, 2)
            mergedList(k, 3) = leftList(i, 3)
            mergedList(k, 4) = leftList(i, 4)
            i = i + 1
        Else
            mergedList(k, 1) = rightList(j, 1)
            mergedList(k, 2) = rightList(j, 2)
            mergedList(k, 3) = rightList(j, 3)
            mergedList(k, 4) = rightList(j, 4)
            j = j + 1
        End If
        k = k + 1
    Loop
    
    ' Copy any remaining elements from leftList
    Do While i <= leftCount
        mergedList(k, 1) = leftList(i, 1)
        mergedList(k, 2) = leftList(i, 2)
        mergedList(k, 3) = leftList(i, 3)
        mergedList(k, 4) = leftList(i, 4)
        i = i + 1
        k = k + 1
    Loop
    
    ' Copy any remaining elements from rightList
    Do While j <= rightCount
        mergedList(k, 1) = rightList(j, 1)
        mergedList(k, 2) = rightList(j, 2)
        mergedList(k, 3) = rightList(j, 3)
        mergedList(k, 4) = rightList(j, 4)
        j = j + 1
        k = k + 1
    Loop
    
    Debug.Print "Merged List:"
    PrintArray mergedList
    
    Merge = mergedList
End Function
Function list_By_Column() As Collection
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
    
    Set list_By_Column = dataLists
End Function
' Helper function to print 2D arrays
Sub PrintArray(arr As Variant)
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Debug.Print arr(i, 2) ' Print the date element (column 2)
    Next i
    Debug.Print "----"
End Sub
