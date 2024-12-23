Attribute VB_Name = "Module1"
Option Explicit
Sub MergeSortedLists()
    Dim reorderedLists As Collection
    Set reorderedLists = New Collection
    Dim mergedList() As Variant
    Dim heap() As Variant
    Dim heapSize As Long
    Dim i As Long
    Dim resultSize As Long

    ' Initialize heap
    heapSize = 0
    ReDim heap(1 To 21) ' Assuming max 21 lists

    ' Insert the first element from each list into the heap
    For i = 1 To reorderedLists.Count
        If Not IsEmpty(reorderedLists(i)) Then
            InsertIntoHeap heap, heapSize, reorderedLists(i)(1), i, 1 ' Insert the first element of each list
        End If
    Next i

    ' Merging process
    resultSize = 0
    ReDim mergedList(1 To 100) ' Size of final result (will adjust as needed)

    Do While heapSize > 0
        Dim minValue As Variant, listIndex As Long, elementIndex As Long
        ExtractMin heap, heapSize, minValue, listIndex, elementIndex

        ' Add the minimum value to the merged list
        resultSize = resultSize + 1

        ' Check to make sure list doesn't overflow and adds 100 space
        If resultSize > UBound(mergedList) Then
            ReDim Preserve mergedList(1 To UBound(mergedList) + 100)
        End If

        ' Store the minimum value
        mergedList(resultSize) = minValue

        ' Insert the next element from the same list into the heap
        If elementIndex <= UBound(reorderedLists(listIndex)) Then
            InsertIntoHeap heap, heapSize, reorderedLists(listIndex)(elementIndex), listIndex, elementIndex + 1
        End If
    Loop

    ' Trim the merged list to the correct size
    ReDim Preserve mergedList(1 To resultSize)

    ' Output the merged list to the Immediate Window
    For i = 1 To resultSize
        Debug.Print mergedList(i)(0), mergedList(i)(1), mergedList(i)(2) ' Print date, name, and location
    Next i

    ' Prepare output for message box
    Dim outputData As String
    outputData = "Merged List:" & vbCrLf

    For i = 1 To resultSize
        outputData = outputData & mergedList(i)(0) & " | " & mergedList(i)(1) & " | " & mergedList(i)(2) & vbCrLf ' Date | Name | Location
    Next i

    ' Show the merged data in a message box
    MsgBox outputData, vbInformation, "Merged Sorted Lists"
End Sub
Sub Create_2D_List()
    Dim ws As Worksheet
    Dim dataLists As New Collection
    Dim colNum As Long
    Dim outputData As String
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
        'MsgBox outputData, vbInformation, "Bouldering Data - List " & j
    Next j
End Sub

' Insert a value into the heap
Sub InsertIntoHeap(ByRef heap() As Variant, ByRef heapSize As Long, ByVal value As Variant, ByVal listIndex As Long, ByVal elementIndex As Long)
    heapSize = heapSize + 1
    ReDim Preserve heap(1 To heapSize)

    Dim i As Long
    i = heapSize

    ' Bubble up the value to maintain the min-heap property
    Do While i > 1 And heap(i \ 2)(0) > value(0) ' Compare by date
        heap(i) = heap(i \ 2)
        i = i \ 2
    Loop

    ' Insert the value at its correct position
    heap(i) = Array(value(0), value(1), value(2), listIndex, elementIndex) ' Store the value along with its list index and element index
End Sub

' Extract the minimum value from the heap
Sub ExtractMin(ByRef heap() As Variant, ByRef heapSize As Long, ByRef minValue As Variant, ByRef listIndex As Long, ByRef elementIndex As Long)
    minValue = heap(1)(0) ' First element is the date
    listIndex = heap(1)(3) ' List index
    elementIndex = heap(1)(4) ' Element index

    ' Move the last element to the root and bubble it down
    heap(1) = heap(heapSize)
    heapSize = heapSize - 1
    ReDim Preserve heap(1 To heapSize)

    Dim i As Long, child As Long
    i = 1

    Do While i * 2 <= heapSize
        child = i * 2
        If child < heapSize And heap(child)(0) > heap(child + 1)(0) Then ' Compare dates
            child = child + 1
        End If
        If heap(i)(0) <= heap(child)(0) Then Exit Do
        ' Swap the values
        Dim temp As Variant
        temp = heap(i)
        heap(i) = heap(child)
        heap(child) = temp
        i = child
    Loop
End Sub
