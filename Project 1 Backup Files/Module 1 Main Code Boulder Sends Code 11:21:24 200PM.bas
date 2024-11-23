Attribute VB_Name = "Module1"
Public var_Sorted As Variant ' store sorted list so we can reference it for then sort dropdown
Public coll_Sorted As Collection
' if sorted by emoji store here
Public emoji_Sorted
Public custom_Sorted

Public numCol As Long ' flexible columns of data
Public startRow As Long ' flexible startRow of data

Public data_ws As Worksheet ' reference to which worksheet we pull from either "Send Data" or "Project Data"
Public output As Worksheet ' for reference when we use the tab "Sort"
Sub custom_Grade(v_Start As String, v_End As String)
    Dim current_List As Variant
    
     If Not IsEmpty(emoji_Sorted) Then
        current_List = emoji_Sorted
    ElseIf Not IsEmpty(var_Sorted) Then
         current_List = var_Sorted
    Else
        MsgBox ("Sort first by 'V-grade', 'Date', or 'Location'")
        Exit Sub
    End If
    
    Dim temp_Var As Variant
    ' so my goal is to iterate once through the variant and make a new variant to store our filtered values
    Dim v_Start_Before As String
    Dim v_End_After As String
    
    Dim v_N_Start, v_N_End_Num, temp_Num As Integer
    
    Dim i As Long
    
    Dim vgradekeys As Variant
    
    vgradekeys = Array("4c/V0", "5a/V1", "5b/V1", "5c/V2", "6a/V3", "6a+/V3", _
    "6b/V4", "6b+/V4", "6c/V5", "6c+/V5", "7a/V6", "7a+/V7", "7b/V8", "7b+/V8", _
    "7c/V9", "7c+/V10", "8a/V11", "8a+/V12", "8b/V13", "8b+/V14", "8c/V15")
        
    For i = 0 To UBound(vgradekeys) - 1
   
        If v_Start = vgradekeys(i) And i <> 0 Then
             v_Start_Before = vgradekeys(i - 1)
        End If
        
        If v_End = vgradekeys(i) And i <> UBound(vgradekeys) - 1 Then
            v_End_After = vgradekeys(i + 1)
        End If
        
    Next i
    
                       
    Dim v_temp As String
    
    v_Start_Num = ExtractEndNumber(v_Start)
    v_End_Num = ExtractEndNumber(v_End)
    
    ' set new list as the same size and will remove blanks at the end
    ReDim temp_Var(1 To UBound(current_List), 1 To numCol)
    
    'to efficiently sort we need to iterate once and perform one check for O(n) time
    'due to having multiple V grades for each euro grade we much check the
    'we must check the value before / after and not include it
    
    ' add to list if between the two grades
    For i = 1 To UBound(current_List)
        'extract end number
        ' pass the v-grade as string
        v_temp = current_List(i, 1)
        temp_Num = ExtractEndNumber(v_temp)
        ' check if V grade low <= its grade <= v grade high, and its not the europe grade b4/ after
        If v_Start_Num <= temp_Num And temp_Num <= v_End_Num _
        And current_List(i, 1) <> v_Start_Before And current_List(i, 1) <> v_End_After Then
            temp_Var(i, 1) = current_List(i, 1) ' Column A
            temp_Var(i, 2) = current_List(i, 2) ' Column B
            temp_Var(i, 3) = current_List(i, 3) ' Column C
            temp_Var(i, 4) = current_List(i, 4) ' Column D
        End If
    Next i
    ' remove blanks
    temp_Var = RemoveBlanks(temp_Var)
    'remove prior contents
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents
    'output filtered list
    output.Range("A" & startRow).Resize(UBound(temp_Var), numCol).Value = temp_Var
    custom_Sorted = temp_Var

    Call borders
End Sub
Sub ascend()
    If Not IsEmpty(custom_Sorted) Then
        output.Range("A" & startRow).Resize(UBound(custom_Sorted, 1), numCol).Value = custom_Sorted
        Exit Sub
    ElseIf Not IsEmpty(emoji_Sorted) Then
        output.Range("A" & startRow).Resize(UBound(emoji_Sorted, 1), numCol).Value = emoji_Sorted
        Exit Sub
    ElseIf Not IsEmpty(var_Sorted) Then
        output.Range("A" & startRow).Resize(UBound(var_Sorted, 1), numCol).Value = var_Sorted
        Exit Sub
    Else
        MsgBox ("Sort first by 'V-grade', 'Date', or 'Location'")
    End If
End Sub
Sub descend()
    Dim flipped_List As Variant
    If Not IsEmpty(custom_Sorted) Then
        flipped_List = FlipArray(custom_Sorted)
        output.Range("A" & startRow).Resize(UBound(flipped_List, 1), numCol).Value = flipped_List
        Exit Sub
     ElseIf Not IsEmpty(emoji_Sorted) Then
        ' flip list
        flipped_List = FlipArray(emoji_Sorted)
        ' output list
        output.Range("A" & startRow).Resize(UBound(flipped_List, 1), numCol).Value = flipped_List
        Exit Sub
    ElseIf Not IsEmpty(var_Sorted) Then
        flipped_List = FlipArray(var_Sorted)
        output.Range("A" & startRow).Resize(UBound(flipped_List, 1), numCol).Value = flipped_List
        Exit Sub
    Else
        MsgBox ("Sort first by V-grade or Date")
    End If
    
End Sub
Function FlipArray(inputArray As Variant) As Variant
    Dim flippedArray As Variant
    Dim i As Long, j As Long
    Dim numRows As Long, numCols As Long
    ' Ensure numCol is initialized
    If numCol = 0 Then
        MsgBox "Error: numCol is not initialized."
        Exit Function
    End If

    ' Get the number of rows and columns in the input array
    numRows = UBound(inputArray, 1)

    ' Initialize the flipped array with the same dimensions
    ReDim flippedArray(1 To numRows, 1 To numCol)

    ' Loop through the input array from the last row to the first
    For i = 1 To numRows
        For j = 1 To numCol
            ' Copy the values in reverse order from inputArray to flippedArray
            flippedArray(i, j) = inputArray(numRows - i + 1, j)
        Next j
    Next i

    ' Return the flipped array
    FlipArray = flippedArray
End Function
Sub Project_Sort()
    If IsEmpty(var_Sorted) Then
        MsgBox ("data was cleared, please re-sort data")
        Exit Sub
    
    Else
        tempVar = var_Sorted
    End If
    
    ' Find the last row in the data
    Dim lastRow As Long
    lastRow = UBound(var_Sorted, 1)
    ' Initialize an array for filtered data
    Dim dataList() As Variant
    ReDim dataList(1 To lastRow, 1 To numCol)

    ' Filter data based on emoji
    Dim t As Long
    t = 1 'temp variable
    
    Dim i As Long
    For i = 1 To UBound(var_Sorted, 1)
        ' check if no emojis are in the climb name and check if it's not empty
        If (InStr(var_Sorted(i, 3), output.Cells(1, 8)) = 0) And (InStr(var_Sorted(i, 3), output.Cells(1, 9)) = 0) _
        And (InStr(var_Sorted(i, 3), output.Cells(1, 10)) = 0) And var_Sorted(i, 2) <> "" Then
            dataList(t, 1) = var_Sorted(i, 1)  ' Grade
            dataList(t, 2) = var_Sorted(i, 2)  ' Date
            dataList(t, 3) = var_Sorted(i, 3)  ' Name
            dataList(t, 4) = var_Sorted(i, 4)  ' Location
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
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents
    output.Range("A" & startRow).Resize(UBound(dataList, 1), numCol).Value = dataList
    
    emoji_Sorted = RemoveBlanks(dataList) 'store the most recent affected list here
    
End Sub
Sub OS_F_OneS_Sort(emoji As String)
    Dim tempVar As Variant
    
    If IsEmpty(var_Sorted) Then
        MsgBox ("data was cleared, please re-sort data")
        Exit Sub
    
    Else
        tempVar = var_Sorted
    End If
        ' Find the last row in the data
        Dim lastRow As Long
        lastRow = UBound(tempVar, 1)
        ' Initialize an array for filtered data
        Dim dataList() As Variant
        ReDim dataList(1 To lastRow, 1 To numCol)
    
        ' Filter data based on emoji
        Dim t As Long
        t = 1 'temp variable
        
        Dim i As Long
        For i = 1 To UBound(tempVar, 1)
            If InStr(tempVar(i, 3), emoji) <> 0 Then
                dataList(t, 1) = tempVar(i, 1)  ' Grade
                dataList(t, 2) = tempVar(i, 2)  ' Date
                dataList(t, 3) = tempVar(i, 3)  ' Name
                dataList(t, 4) = tempVar(i, 4)  ' Location
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
        output.Rows(startRow & ":" & output.Rows.Count).ClearContents
        output.Range("A" & startRow).Resize(UBound(dataList, 1), numCol).Value = dataList
        
        emoji_Sorted = RemoveBlanks(dataList) 'store the most recent affected list here
    
End Sub
Sub mergeSort_By_Date()

    ' reset our public variable for collection sorted
    Set coll_Sorted = New Collection
    ' this is a temp variant to add to coll_Sorted
    Dim tempVar As Variant
    Dim dataLists As New Collection
    Dim sortedList As Variant
    Dim lastRow As Long
    
    ' run sort by columns in the data
    Set dataLists = climbData()
    
    ' Recursive merge on reorderedLists
    sortedList = RecursiveMerge(dataLists, 1, dataLists.Count)
     
   ' Get the total number of rows in sortedList
    lastRow = UBound(sortedList, 1)
    
    ' use our own blank function to iterate through the list because we have to update coll_sorted anyways
    
    ' Step 1: Count the number of blank rows at the start
    Dim blankCount As Long
    blankCount = 0
    Dim i As Long
    i = 1
    Do While sortedList(i, 2) = ""
        blankCount = blankCount + 1
        i = i + 1
    Loop
    
    ' Step 2: Shift the non-blank rows up (remove the blanks) and make the new collection
    
    Dim j As Long
    k = 1 ' j will be used to shift non-blank rows to the front since we know they will be in the front
   
   Dim currentdate As Variant
   Dim sRow, endrow, numRows As Long
   
   ' intialize variables
    sRow = blankCount + 1
   
   currentdate = sortedList(blankCount + 1, 2)
   
    For i = blankCount + 1 To lastRow
        sortedList(k, 1) = sortedList(i, 1)
        sortedList(k, 2) = sortedList(i, 2)
        sortedList(k, 3) = sortedList(i, 3)
        sortedList(k, 4) = sortedList(i, 4)
        k = k + 1
        
        If sortedList(i, 2) <> currentdate Then
            ' Date changed, add the previous group to the collection
            endrow = i - 1
            numRows = endrow - sRow + 1
            If lastRow > 0 Then
                ' Create a temporary array for the group
                ReDim tempVar(1 To numRows, 1 To 4)
                For j = 1 To numRows
                    tempVar(j, 1) = sortedList(sRow + j - 1, 1)
                    tempVar(j, 2) = sortedList(sRow + j - 1, 2)
                    tempVar(j, 3) = sortedList(sRow + j - 1, 3)
                    tempVar(j, 4) = sortedList(sRow + j - 1, 4)
                Next j

                ' Add the grouped array to the collection
                coll_Sorted.Add tempVar
            End If

            ' Update the start row and current date
            sRow = i
            currentdate = sortedList(i, 2)
        End If
        
    Next i
    
    ' Add the last group to the collection
    numRows = lastRow - sRow + 1
    If numRows > 0 Then
        ReDim tempVar(1 To numRows, 1 To 4)
        For j = 1 To numRows
            tempVar(j, 1) = sortedList(sRow + j - 1, 1)
            tempVar(j, 2) = sortedList(sRow + j - 1, 2)
            tempVar(j, 3) = sortedList(sRow + j - 1, 3)
            tempVar(j, 4) = sortedList(sRow + j - 1, 4)
        Next j
        coll_Sorted.Add tempVar
    End If
    
    ' Step 3: Add blank rows to the back
    For i = k To lastRow
        sortedList(i, 1) = ""
        sortedList(i, 2) = ""
        sortedList(i, 3) = ""
        sortedList(i, 4) = ""
    Next i

    ' Now sortedList has the blank rows moved to the back.

    ' Clear previous data
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents ' Clear only rows below row 3

    ' Resize the range and assign sortedList values to the worksheet w/ blanks
    'output.Range("A3").Resize(lastRow, numCol).value = sortedList
    
    'to not have the blank values
    output.Range("A" & startRow).Resize(lastRow - blankCount, numCol).Value = sortedList
    
    var_Sorted = output.Range("A" & startRow).Resize(lastRow - blankCount, numCol).Value ' store the value without blanks back to ascend and descend
    
End Sub
Function RecursiveMerge(lists As Collection, left As Long, right As Long) As Variant
    If left = right Then
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
    
    ' old function
    'Set dataLists = list_By_Column() 'Save list by column
    ' new function
    Set dataLists = climbData()
    
    Set reorderedLists = list_By_Row(dataLists) 'Save list by row
    
    'output.Cells.Clear ' Clear any previous data on Sheet2
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents ' Clear only rows below row 5
    
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
    output.Range("A" & startRow).Resize(rowCount, numCol).Value = sort_By_V_Grade
    
    var_Sorted = RemoveBlanks(sort_By_V_Grade)
    
End Sub
Sub list_By_Rows()

    Dim dataLists As New Collection
    Dim reorderedLists As New Collection 'This is the collection of data we are outputting
    
    ' old function
    'Set dataLists = list_By_Column() 'Save list by column
    ' new function
    Set dataLists = climbData()
    
    
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
    Set coll_Sorted = reorderedLists
End Function
Function RemoveBlanks(inputArray As Variant) As Variant
    Dim i As Long, j As Long
    Dim temparray As Variant
    Dim nonBlankCount As Long
    
    ' Count the non-blank rows
    For i = LBound(inputArray, 1) To UBound(inputArray, 1)
        If inputArray(i, 1) <> "" Then
            nonBlankCount = nonBlankCount + 1
        End If
    Next i

   ' If there are no blanks, just return the original array
    If nonBlankCount = (UBound(inputArray, 1) - LBound(inputArray, 1) + 1) Then
        RemoveBlanks = inputArray
        Exit Function
    End If
    '
    If nonBlankCount = 0 Then
        RemoveBlanks = inputArray
        Exit Function
    End If
    

    ' Initialize tempArray with the correct dimensions
    ReDim temparray(1 To nonBlankCount, 1 To numCol)

    ' Populate tempArray with non-blank rows from inputArray
    Dim newRow As Long
    newRow = 1
    For i = LBound(inputArray, 1) To UBound(inputArray, 1)
        If inputArray(i, 1) <> "" Then
            For j = 1 To numCol
                temparray(newRow, j) = inputArray(i, j)
            Next j
            newRow = newRow + 1
        End If
    Next i

    ' Return the array without blanks
    RemoveBlanks = temparray
End Function
Function climbData() As Collection
    ' ***MAIN FUNCTION PULLS FROM DATA TABS***
    ' set public variables
    
    numCol = 4 'This a public variable where we have 4 columns of data
    startRow = 5 'This a public variable where we are starting on the 5th row
    var_Then_Sorted = Empty 'reset so if this is run (when making sort by date or sort by V-Grade)
    emoji_Sorted = Empty
    'so that ascend/descend sorting works properly
    Set output = ThisWorkbook.Worksheets("Sort")
    Dim current_ws As String
      
    ' set current worksheet to pull data from
    current_ws = output.Cells(2, 1)
    Set data_ws = ThisWorkbook.Worksheets(current_ws)
    'End If
    
    Dim dataLists As Collection 'collection we are going to be adding 2d arrays variant(1 to {1 to 20},1 to 4)
    Set dataLists = New Collection
    
    'this is the 2d array wewill be adding
    Dim gradeData() As Variant 'use gradeData() to dim a specificed size, gradeData is better for ranges
    
    Dim currentRow, t_data, i, j, k As Long
    Dim currentGradePos As Range ' variable for keeping track of last V-Grade
    Dim dataRange As Range
    Dim temparray As Variant ' used to add the range of values say B2 to D5 and then add the fixed values of currentgradepos
     ' row A
    k = 1
     ' iterate through the columns A, E, I that have the V-Grade Data Respectively
    Do While k < 10
        i = 1 ' for inner loop
        currentRow = 4 ' data starts at row 4
        ' t_data only iterates for data w/o blanks, currentRow keeps iterating with data and blanks
        t_data = currentRow
        ' set at first v-grade
        Set currentGradePos = data_ws.Cells(currentRow, k)
        
        ' loop through rows until 7 arrays are added to the collection
        Do While i < 8
            ' check if column k has a "V" grade
            If data_ws.Cells(currentRow, k).Value Like "*V*" Then
                ' if V is found and it is not the start value then add it to the array
                If currentRow > 4 Then
                    ' define the range say B2:D6
                    Set dataRange = data_ws.Range(data_ws.Cells(currentGradePos.Row, k + 1), data_ws.Cells(t_data, numCol + k - 1))
                
                    ' load the range data directly into a temporary array
                    temparray = dataRange.Value
                    
                    ' set the right size for gradeData
                    ReDim gradeData(1 To t_data + 1 - currentGradePos.Row, 1 To numCol)
                    ' set the fixed value for the first column (say 1 to 5, 1)
                    For j = 1 To t_data + 1 - currentGradePos.Row
                        gradeData(j, 1) = currentGradePos.Value
                    Next j
                
                    ' copy the range data (B2:D6) into the rest of the array (1 to {amount of rows with data}, 2 to {numcol})
                    For j = 1 To t_data + 1 - currentGradePos.Row
                        gradeData(j, 2) = temparray(j, 1) ' Column B
                        gradeData(j, 3) = temparray(j, 2) ' Column C
                        gradeData(j, 4) = temparray(j, 3) ' Column D
                    Next j
                    
                    dataLists.Add gradeData
                
                    ' reset currentgradepos for the new V value
                    Set currentGradePos = data_ws.Cells(currentRow, k)
                    
                    t_data = currentRow ' reset the temp data storing the start of the row with data
                    
                    i = i + 1 ' increment the loop
                End If
           ' if v is not found, check if there's data in columns B
            ElseIf data_ws.Cells(currentRow, k + 1).Value <> "" Then
                t_data = t_data + 1
            End If
                
        ' increment currentrow which is keeping track of what cell we are on
        currentRow = currentRow + 1
        Loop
    
    k = k + 4 ' k will be 1, 5, 9 respectively
    
    Loop
    
    ' return function
    Set climbData = dataLists
    
End Function
Function location_Sort() As Collection
    Dim dataLists As New Collection
    Set dataLists = climbData()
    
     'reset recursive merge sort if var_sorted is empty ; more efficient to sort by date then by location
    var_Sorted = RecursiveMerge(dataLists, 1, dataLists.Count)

    
    '**************MAY EDIT maybe I can keep track of it is it run and just use the preexisting var_sorted value instead of running it again***
    'If IsEmpty(var_Sorted) Then
        'Set dataLists = climbData()
        'var_Sorted = RecursiveMerge(dataLists, 1, dataLists.Count)
    
    Set dataLists = New Collection
    
    'this is the 2d array wewill be adding
    Dim locData() As Variant 'use gradeData() to dim a specificed size, gradeData is better for ranges
    
    Dim startNum, i, j As Long
    Dim currentLoc As String ' variable for keeping track of last V-Grade
    Dim dataRange As Range
    Dim temparray As Variant ' used to add the range of values say B2 to D5 and then add the fixed values of currentgradepos
    ' start index at first row of list
    i = 1
     ' set a variable to keep track of start and name of first location
    startNum = 1
    currentLoc = var_Sorted(i, numCol)
    
    ' loop through rows until there's no data
    Do While i < UBound(var_Sorted, 1) + 1
        ' check if column k has a " grade
        If var_Sorted(i, numCol) <> currentLoc Then
            ' if V is found and it is not the start value then add it to the array
            ' define the range say B2:D6
            
            ' set the right size for gradeData
            ReDim locData(1 To i - startNum, 1 To numCol)

            ' copy the range data (B2:D6) into the rest of the array (1 to {amount of rows with data}, 2 to {numcol})
            For j = 1 To i - startNum
                locData(j, 1) = var_Sorted(startNum + j - 1, 1)  ' Column B
                locData(j, 2) = var_Sorted(startNum + j - 1, 2) ' Column B
                locData(j, 3) = var_Sorted(startNum + j - 1, 3) ' Column C
                locData(j, 4) = var_Sorted(startNum + j - 1, 4) ' Column D`
            Next j
            
            dataLists.Add locData
        
            ' reset currentgradepos for the new V value
            currentLoc = var_Sorted(i, numCol)
            startNum = i
        End If
        
        i = i + 1 ' increment the loop
            
    Loop
    
    ' set the right size for locData
    ReDim locData(1 To i - startNum, 1 To numCol)

    ' copy the range data (B2:D6) into the rest of the array (1 to {amount of rows with data}, 2 to {numcol})
    For j = 1 To i - startNum
        locData(j, 1) = var_Sorted(startNum + j - 1, 1)  ' Column B
        locData(j, 2) = var_Sorted(startNum + j - 1, 2) ' Column B
        locData(j, 3) = var_Sorted(startNum + j - 1, 3) ' Column C
        locData(j, 4) = var_Sorted(startNum + j - 1, 4) ' Column D`
    Next j
    
    dataLists.Add locData
    
    ' return function
    Set location_Sort = dataLists
    
End Function
Function loc_Sort_Date(reorderedLists As Collection) As Collection
    Dim mergedData As Collection
    Dim i As Long, j As Long
    Dim currentData As Variant, key As String
    Dim mergedArray As Variant
    Dim keyExists As Boolean

    ' Initialize the collection
    Set mergedData = New Collection

    ' Iterate through reorderedLists
    For i = 1 To reorderedLists.Count
        currentData = reorderedLists(i)
        key = CStr(currentData(1, 4))

        ' Check if the key already exists in mergedData
        keyExists = False
        For j = 1 To mergedData.Count
            If mergedData(j)(1, 4) = key Then
                mergedArray = Merge(mergedData(j), currentData)
                mergedData.Remove j
                mergedData.Add mergedArray
                keyExists = True
                Exit For
            End If
        Next j

        ' If the key does not exist, add the current data
        If Not keyExists Then
            mergedData.Add currentData
        End If
    Next i

    ' Return the merged collection
    Set loc_Sort_Date = mergedData
    ' set public as the recent merged collection
    Set coll_Sorted = mergedData
End Function
Sub loc_Sort_By_Date_Ouput()
    Dim reorderedLists As New Collection
    Set reorderedLists = location_Sort()
    Set reorderedLists = loc_Sort_Date(reorderedLists)
    
    loc_Sort_Output reorderedLists
End Sub
Function loc_Sort_Output(reorderedLists As Collection)
    
    'output for sort by location sort
    
   Dim sort_By_Loc As Variant
    
    'output.Cells.Clear ' Clear any previous data on Sheet2
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents ' Clear only rows below row 5
    
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
    ReDim sort_By_Loc(1 To rowCount, 1 To 4) ' Assuming 4 columns in each array
    ' Loop through sortedList and write values to sorted sheet
    For i = 1 To reorderedLists.Count
        currentArray = reorderedLists(i)
        ' Access the bounds of the current array

        Dim upperBound As Long
        upperBound = UBound(currentArray)

        ' Loop through the elements of the current array
        For j = 1 To upperBound ' basically 1 to max
            If currentArray(j, 2) <> "" Then ' get rid of blanks here
                sort_By_Loc(t, 1) = currentArray(j, 1) ' Column A
                sort_By_Loc(t, 2) = currentArray(j, 2) ' Column B
                sort_By_Loc(t, 3) = currentArray(j, 3) ' Column C
                sort_By_Loc(t, 4) = currentArray(j, 4) ' Column D
                t = t + 1
            End If
        Next j
    Next i
    
    ' Write the entire array to the worksheet at once, starting from A3
    output.Range("A" & startRow).Resize(rowCount, numCol).Value = sort_By_Loc
    
    var_Sorted = RemoveBlanks(sort_By_Loc)
    
End Function
Sub then_VGrade_Sort()
    ' this would reset the location sort
    'Dim reorderedLists As New Collection
    'Set reorderedLists = location_Sort()
    'Set reorderedLists = loc_Sort_Date(reorderedLists)
    
    If Not coll_Sorted Is Nothing Then
        Dim output As Worksheet
        Set output = ThisWorkbook.Sheets("Sort")
        
        Dim loc_VGrade As New Collection
        
        ' Get the current variant (which should be an array)
        Dim currentArray As Variant
        Dim temparray As Variant
        
        Dim i As Long
        ' iterate through the variants in the loc_sort collection and add new variants to our loc_vgrade collection
        For i = 1 To coll_Sorted.Count
        
            currentArray = coll_Sorted(i)
            
            temparray = var_VGrade_By_Loc(currentArray)
            
            loc_VGrade.Add temparray
        Next i
    
    ' call output function
    loc_Sort_Output loc_VGrade
    
    Else
        MsgBox ("Please sort by location first")
    End If
End Sub
Function var_VGrade_By_Loc(currentArray As Variant) As Variant
    ' set our variant we are returning to same length
    Dim temp_array As Variant
    ReDim temp_array(1 To UBound(currentArray), 1 To numCol)
    
    Dim temp_Var As Variant
    ' so my goal is to iterate through the location sort for each location (2d variant, temp array,
    'we are going to make individual collections for every v-grade and add them chronolocially back to the variant
    Dim i, j As Long
    Dim vCollections(1 To 22) As Collection
    Dim vgradekeys As Variant
     ' Initialize collections for each V-grade
    For i = LBound(vCollections) To UBound(vCollections)
        Set vCollections(i) = New Collection
    Next i
    
    vgradekeys = Array("4c/V0", "5a/V1", "5b/V1", "5c/V2", "6a/V3", "6a+/V3", _
                       "6b/V4", "6b+/V4", "6c/V5", "6c+/V5", "7a/V6", "7a+/V7", _
                       "7b/V8", "7b+/V8", "7c/V9", "7c+/V10", "8a/V11", "8a+/V12", _
                       "8b/V13", "8b+/V14", "8c/V15")
    ' Iterate through the current array and sort by V-grade
    For i = 1 To UBound(currentArray)
        ' Create a temporary array for the current row
        ReDim temp_Var(1 To 1, 1 To numCol)
        temp_Var(1, 1) = currentArray(i, 1) ' Column A
        temp_Var(1, 2) = currentArray(i, 2) ' Column B
        temp_Var(1, 3) = currentArray(i, 3) ' Column C
        temp_Var(1, 4) = currentArray(i, 4) ' Column D
        
        ' Find the index for the current V-grade
        For j = LBound(vgradekeys) To UBound(vgradekeys)
            If currentArray(i, 1) = vgradekeys(j) Then
                vCollections(j + 1).Add temp_Var
                Exit For
            End If
        Next j
        
        ' If no match is found, notify the user
        If j > UBound(vgradekeys) Then
            MsgBox "V-grade not known for: " & currentArray(i, 1)
        End If
    Next i
   
    t = 1
    i = 1
    j = 1
    
    ' Array of V-grade variables
   
    
    ' Loop through each V-grade in the array
    For i = LBound(vgradekeys) To UBound(vgradekeys)
        If Not vCollections(i + 1) Is Nothing Then
            ' Process the current V-grade collection
            For j = 1 To vCollections(i + 1).Count
                temp_Var = vCollections(i + 1)(j)
                temp_array(t, 1) = temp_Var(1, 1) ' Column A
                temp_array(t, 2) = temp_Var(1, 2) ' Column B
                temp_array(t, 3) = temp_Var(1, 3) ' Column C
                temp_array(t, 4) = temp_Var(1, 4) ' Column D
                t = t + 1
            Next j
        End If
    Next i
    var_VGrade_By_Loc = temp_array
End Function
Function ExtractEndNumber(Grade As String) As Integer
    Dim Length As Integer
    Dim LastChar As String
    Dim SecondLastChar As String
    Dim EndNumber As String

    ' Get the length of the input string
    Length = Len(Grade)

    ' Extract the last character
    LastChar = mid(Grade, Length, 1)
    If LastChar Like "[0-9]" Then
        EndNumber = LastChar
    End If

    ' Check the second-to-last character, if it exists
    If Length > 1 Then
        SecondLastChar = mid(Grade, Length - 1, 1)
        If SecondLastChar Like "[0-9]" Then
            EndNumber = SecondLastChar & EndNumber
        End If
    End If

    ' Convert and return the numeric value, defaulting to 0 if no number is found
    If EndNumber <> "" Then
        ExtractEndNumber = CInt(EndNumber)
    Else
        ExtractEndNumber = 0
    End If
End Function
Sub addDataValidation()
    Dim GradeList As String
    
    GradeList = "4c/V0,5a/V1,5b/V1,5c/V2,6a/V3,6a+/V3," & _
                "6b/V4,6b+/V4,6c/V5,6c+/V5,7a/V6,7a+/V7," & _
                "7b/V8,7b+/V8,7c/V9,7c+/V10,8a/V11,8a+/V12," & _
                "8b/V13,8b+/V14,8c/V15"
    
    ' Add data validation to a specific cell
    With output.Range("F3").Validation
        .Delete ' Remove existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=GradeList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    With output.Range("F4").Validation
        .Delete ' Remove existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=GradeList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    MsgBox "Please select start and end V-Grade from cell F3 & F4"
End Sub
Function ValidGrade(Grade As String) As Boolean
    Dim ValidGrades As Variant
    
    ValidGrade = False
    
    Dim i As Integer
    
    ' List of valid grades
    ValidGrades = Array("4c/V0", "5a/V1", "5b/V1", "5c/V2", "6a/V3", "6a+/V3", _
                        "6b/V4", "6b+/V4", "6c/V5", "6c+/V5", "7a/V6", "7a+/V7", _
                        "7b/V8", "7b+/V8", "7c/V9", "7c+/V10", "8a/V11", "8a+/V12", _
                        "8b/V13", "8b+/V14", "8c/V15")
    
    ' Validate the input
    For i = LBound(ValidGrades) To UBound(ValidGrades)
        If Grade = ValidGrades(i) Then
            ValidGrade = True
            Exit For
        End If
    Next i
    
End Function
