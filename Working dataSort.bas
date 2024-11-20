Attribute VB_Name = "Module2"
Sub DebugPrintDataLists()
    Dim dataLists As New Collection
    Set dataLists = ClimbData()
    
    Dim p As Long, l As Long
    Dim gradeData As Variant
    
    ' Loop through each 2D array in the dataLists collection
    For p = 1 To dataLists.Count
        
        ' Retrieve the 2D array for the current grade
        gradeData = dataLists(p)
        Debug.Print "Data for " & gradeData(1, 1) & ":"
        
        ' Loop through the rows of the 2D array (up to 10 rows)
        For l = 1 To UBound(gradeData)
            ' Check if the row is not empty (first column should have the grade)
            If Not IsEmpty(gradeData(l, 1)) Then
                Debug.Print gradeData(l, 1) & vbTab & _
                            gradeData(l, 2) & vbTab & _
                            gradeData(l, 3) & vbTab & _
                            gradeData(l, 4)
            End If
        Next l
        
        Debug.Print String(50, "-") ' Separator line for clarity
    Next p
End Sub
Function ClimbData() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Send Data") ' Adjust as needed
    
    Dim dataLists As Collection 'collectionwe are going to be adding 2d arrays variant(1 to {1 to 20},1 to 4)
    Set dataLists = New Collection
    
    'this is the 2d array wewill be adding
    Dim gradeData() As Variant 'use gradeData() to dim a specificed size, gradeData is better for ranges
    
    Dim colNum As Long
    colNum = 4 ' public variable in other sub
    
    Dim currentRow, t_data, i, j, k As Long
    Dim currentGradePos As Range ' variable for keeping track of last V-Grade
    Dim dataRange As Range
    Dim tempArray As Variant ' used to add the range of values say B2 to D5 and then add the fixed values of currentgradepos
     ' row A
    k = 1
     ' iterate through the columns A, E, I that have the V-Grade Data Respectively
    While k < 10
        i = 1 ' for inner loop
        currentRow = 2 ' data starts at row 2
        ' t_data only iterates for data w/o blanks, currentRow keeps iterating with data and blanks
        t_data = currentRow
        ' set at first v-grade
        Set currentGradePos = ws.Cells(currentRow, k)
        
        ' loop through rows until 7 arrays are added to the collection
        While i < 8
            ' check if column k has a "V" grade
            If ws.Cells(currentRow, k).value Like "*V*" Then
                ' if V is found and it is not the start value then add it to the array
                If currentRow > 2 Then
                    ' define the range say B2:D6
                    Set dataRange = ws.Range(ws.Cells(currentGradePos.row, k + 1), ws.Cells(t_data, colNum + k - 1))
                
                    ' load the range data directly into a temporary array
                    tempArray = dataRange.value
                    
                    ' set the right size for gradeData
                    ReDim gradeData(1 To t_data + 1 - currentGradePos.row, 1 To colNum)
                    ' set the fixed value for the first column (say 1 to 5, 1)
                    For j = 1 To t_data + 1 - currentGradePos.row
                        gradeData(j, 1) = currentGradePos.value
                    Next j
                
                    ' copy the range data (B2:D6) into the rest of the array (1 to {amount of rows with data}, 2 to {colNum})
                    For j = 1 To t_data + 1 - currentGradePos.row
                        gradeData(j, 2) = tempArray(j, 1) ' Column B
                        gradeData(j, 3) = tempArray(j, 2) ' Column C
                        gradeData(j, 4) = tempArray(j, 3) ' Column D
                    Next j
                    
                    dataLists.Add gradeData
                
                    ' reset currentgradepos for the new V value
                    Set currentGradePos = ws.Cells(currentRow, k)
                    
                    t_data = currentRow ' reset the temp data storing the start of the row with data
                    
                    i = i + 1 ' increment the loop
                End If
           ' if v is not found, check if there's data in columns B
            ElseIf ws.Cells(currentRow, k + 1).value <> "" Then
                t_data = t_data + 1
            End If
                
        ' increment currentrow which is keeping track of what cell we are on
        currentRow = currentRow + 1
        Wend
    
    k = k + 4 ' k will be 1, 5, 9 respectively
    
    Wend
    
    ' return function
    Set ClimbData = dataLists
    
End Function
