Attribute VB_Name = "Module2"
Public recent_Sorted_List As Variant
Public numCol
Public startRow
Sub onsight_Sort(emoji As String)
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sort")

    ' Find the last row in the data
    Dim lastRow As Long
    lastRow = UBound(recent_Sorted_List, 1)
    ' Initialize an array for filtered data
    Dim dataList() As Variant
    ReDim dataList(1 To lastRow, 1 To numCol)

    ' Filter data based on emoji
    Dim t As Long
    t = 1 'temp variable
    
    Dim i As Long
    For i = 1 To UBound(recent_Sorted_List, 1)
        If InStr(recent_Sorted_List(i, 3), emoji) <> 0 Then
            dataList(t, 1) = recent_Sorted_List(i, 1)  ' Grade
            dataList(t, 2) = recent_Sorted_List(i, 2)  ' Date
            dataList(t, 3) = recent_Sorted_List(i, 3)  ' Name
            dataList(t, 4) = recent_Sorted_List(i, 4)  ' Location
            t = t + 1
        End If
    Next i

    ' Resize dataList based on matched items
    If t = 1 Then
        ' If no matches, create a single empty row with placeholders
        ReDim dataList(1 To 1, 1 To numCol)
        dataList(1, 1) = "No matches found"  ' Optional message
    End If

    ' Clear previous data and write filtered data to the worksheet
    ws1.Rows(startRow & ":" & ws1.Rows.Count).ClearContents
    ws1.Range("A" & startRow).Resize(UBound(dataList, 1), numCol).value = dataList
End Sub

Sub mergeSort_By_Date()

Dim dataLists As New Collection

Set dataLists = list_By_Column() ' use the sort by column, because it is one less function and we are sorting by date anyways

'Recursive merge on reorderedLists
    Dim sortedList As Variant
    sortedList = RecursiveMerge(dataLists, 1, dataLists.Count)
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sort")
    
    lastRow As Long
    'startRow = 3 don't need this variable because i'm copying the data at A3
    ' sorted list has the blanks in the front (for say no V11s sent doesn't have data)
    'next 3 steps are just to move the blanks to the back of the list
   ' Get the total number of rows in sortedList
    lastRow = UBound(sortedList, 1)
    
    ' Step 1: Count the number of blank rows at the start
    Dim blankCount As Long
    blankCount = 0
    Dim i As Long
    i = 1
    Do While sortedList(i, 2) = ""
        blankCount = blankCount + 1
        i = i + 1
    Loop
    ' Step 2: Shift the non-blank rows up (remove the blanks)
    Dim j As Long
    j = 1 ' j will be used to shift non-blank rows to the front
    For i = blankCount + 1 To lastRow
        sortedList(j, 1) = sortedList(i, 1)
        sortedList(j, 2) = sortedList(i, 2)
        sortedList(j, 3) = sortedList(i, 3)
        sortedList(j, 4) = sortedList(i, 4)
        j = j + 1
    Next i
    
    ' Step 3: Add blank rows to the back
    For i = j To lastRow
        sortedList(i, 1) = ""
        sortedList(i, 2) = ""
        sortedList(i, 3) = ""
        sortedList(i, 4) = ""
    Next i

    ' Now sortedList has the blank rows moved to the back.

    ' Clear previous data
    ws1.Rows(startRow & ":" & ws1.Rows.Count).ClearContents ' Clear only rows below row 3

    ' Resize the range and assign sortedList values to the worksheet w/ blanks
    'ws1.Range("A3").Resize(lastRow, numCol).value = sortedList
    
    'to not have the blank values
    ws1.Range("A" & startRow).Resize(lastRow - blankCount, numCol).value = sortedList
    
    recent_Sorted_List = sortedList 'store public variable to be able to edit list
    
    MsgBox "Successfully Sorted Climbs By Date", vbInformation
    
    'used to add data cell by cell which is less effficent and doesn't have blank handling
    'if I want to put in a message for the V-Grades I have not sent
    
    'dim l as long, t as long
    ' loop through sortedList and write values to Sorted sheet
    'For l = LBound(sortedList, 1) To UBound(sortedList, 1)
        'If sortedList(l, 1) <> "" Then
            'ws1.Cells(t, 1).value = sortedList(l, 1) ' Column A
            'ws1.Cells(t, 2).value = sortedList(l, 2) ' Column B
            'ws1.Cells(t, 3).value = sortedList(l, 3) ' Column C
           ' ws1.Cells(t, 4).value = sortedList(l, 4) ' Column D
            't = t + 1
        'Else
        
        'End If
        
    'Next l
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
Sub lists_By_Columns()
    Set dataLists = list_By_Column
    
    outputData = "Bouldering Sends Summary:" & vbCrLf
 
    ' Loop through each dataList in the collection
    For j = 1 To dataLists.Count
        dataList = dataLists(j) ' Get the dataList from the collection
        
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
Sub sort_By_Rows()
    Dim dataLists As New Collection
    Dim reorderedLists As New Collection 'This is the collection of data we are outputting
    Dim sort_By_V_Grade As Variant
    
    Set dataLists = list_By_Column() 'Save list by column
    
    Set reorderedLists = list_By_Row(dataLists) 'Save list by row
    
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sort")
    
    'ws1.Cells.Clear ' Clear any previous data on Sheet2
    ws1.Rows(startRow & ":" & ws1.Rows.Count).ClearContents ' Clear only rows below row 5
    
        ' Get the current variant (which should be an array)
        Dim currentArray As Variant
        
        Dim c As Long
        c = 1
        
        ' First, count the number of rows and columns
        Dim rowCount As Long
        rowCount = 0
        For c = 1 To reorderedLists.Count
            currentArray = reorderedLists(c)
            rowCount = rowCount + UBound(currentArray, 1)
        Next c
        
    Dim t As Long
    Dim j As Long
    j = 1
    t = 1
    
    ' Resize the output array to fit the data
        ReDim sort_By_V_Grade(1 To rowCount, 1 To 4) ' Assuming 4 columns in each array
    ' Loop through sortedList and write values to sorted sheet
    For i = 1 To reorderedLists.Count
        currentArray = reorderedLists(i)
        ' Access the bounds of the current array

        Dim upperBound As Long
        upperBound = UBound(currentArray)

        ' Loop through the elements of the current array
        For j = 1 To upperBound ' basically 1 to max
            If currentArray(j, 1) <> "" Then ' get rid of blanks here
                sort_By_V_Grade(t, 1) = currentArray(j, 1) ' Column A
                sort_By_V_Grade(t, 2) = currentArray(j, 2) ' Column B
                sort_By_V_Grade(t, 3) = currentArray(j, 3) ' Column C
                sort_By_V_Grade(t, 4) = currentArray(j, 4) ' Column D
                t = t + 1
            End If
        Next j
    Next i
    
    ' Write the entire array to the worksheet at once, starting from A3
    ws1.Range("A" & startRow).Resize(rowCount, numCol).value = sort_By_V_Grade
    
    recent_Sorted_List = sort_By_V_Grade
    
    MsgBox "Successfully Sorted Climbs By V-Grade", vbInformation
End Sub
Sub list_By_Rows()

    Dim dataLists As New Collection
    Dim reorderedLists As New Collection 'This is the collection of data we are outputting
    
    Set dataLists = list_By_Column() 'Save list by column
    
    Set reorderedLists = list_By_Row(dataLists) 'Save list by row
    
    'this is just output data iterating through the collection
    
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
Function list_By_Column() As Collection
    Dim ws As Worksheet
    Dim dataLists As New Collection
    Dim t_col As Long 'temp column
    Dim t_startRow As Long 'temp starting row that resets
    Dim i, j, k, totalDataRows, currentRow, endStartRow, nextStartRow As Long
    Dim bool As Boolean
    
    ' set the worksheet
    Set ws = ThisWorkbook.Sheets("Send Data")
    
    numCol = 4 'This a public variable where we have 4 columns of data
    startRow = 5 'This a public variable where we are starting on the 5th row
    
    t_col = 1 ' this is going to refer to column A then go to column E and I respectively

    ' Initialize variables
    totalDataRows = 0
    currentRow = 0
    endStartRow = 0
    nextStartRow = 2

    t_startRow = 2 ' other variables declared everytime

    For j = 1 To 3
        For k = 1 To 7
            
            Dim dataList() As Variant
        
            totalDataRows = 0 ' reset variables back to the starting position
            
            t_startRow = nextStartRow  ' Start from the first row of data
            
            endStartRow = t_startRow
            
            If ws.Cells(endStartRow, t_col + 1).value = "" Then
                endStartRow = endStartRow + 1 'makes a one value blank list otherwise
                ' make it so it still finds the next starting value for when V shows up if blank
             ' not blank start from the top
            End If
            
            
            bool = True
        
            Do While bool And ws.Cells(endStartRow, t_col + 1).value <> ""
                If InStr(ws.Cells(endStartRow + 1, t_col).value, "V") <> 0 Then
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
            Do While InStr(ws.Cells(currentRow, t_col).value, "V") = 0 And nextStartRow <> endStartRow
                currentRow = currentRow + 1
            Loop
            
            If currentRow <> endStartRow Then
                nextStartRow = currentRow
            End If
            ' in theory nextStartRow should equal currentRow + 1
            
            'for the list from startRow to endStartRow
            
            'then new list starting from nextStartRow
            
            ' Resize the dataList array
            totalDataRows = endStartRow - t_startRow
            currentRow = t_startRow
            
            ReDim dataList(1 To totalDataRows, 1 To numCol) ' numCols including 4 grade, date, name, location
            If totalDataRows >= 1 And ws.Cells(currentRow, colNum + 1).value <> "" Then
        
                ' Fill the dataList array
                For i = 1 To totalDataRows
                    dataList(i, 1) = ws.Cells(t_startRow, t_col).value
                    dataList(i, 2) = ws.Cells(currentRow, t_col + 1).value ' Dates from column B
                    dataList(i, 3) = ws.Cells(currentRow, t_col + 2).value ' Names from column C
                    dataList(i, 4) = ws.Cells(currentRow, t_col + 3).value ' Locations from column D
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
    
    t_col = t_col + 4 ' moving from Column A looking for V to Column E
    
    'reset variables to the top
    totalDataRows = 0
    currentRow = 0
    endStartRow = 0
    nextStartRow = 2
    t_startRow = 2 ' starting at 2 baby
    
    Next j
    
    Set list_By_Column = dataLists
End Function
Function list_By_Row(dataLists As Collection)
    
    Dim reorderedLists As New Collection
    
    ' Add lists to the reordered collection based on the desired sequence
    For i = 1 To 7
        reorderedLists.Add dataLists(i) ' 1
        reorderedLists.Add dataLists(i + 7) ' 8
        reorderedLists.Add dataLists(i + 14) ' 15
    Next i
    ' Now you can use reorderedLists as needed
    
    Set list_By_Row = reorderedLists
End Function

