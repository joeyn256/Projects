Attribute VB_Name = "Module2"
Sub Clean_Up_Projects()
    Clean_Up ThisWorkbook.Worksheets("Project Data")
End Sub

Sub Clean_Up_Sends()
    Clean_Up ThisWorkbook.Worksheets("Send Data")
End Sub
Function Clean_Up(ws As Worksheet)

    numCol = 4 'this a public variable where we have 4 columns of data
    
    Dim currentRow, t_data, i, k, l As Long
    Dim currentGradePos As Range ' variable for keeping track of last V-Grade
    Dim insertRange As Range
    Dim borderRange As Range
    
    ' clear all borders
    ws.Cells.borders.LineStyle = xlNone
    
    ' add the borders to the headings
    l = 1
    Do While l < 10
        Set borderRange = Range(ws.Cells(3, l), ws.Cells(3, l + numCol - 1))
        AddBorder borderRange
        l = l + 4
    Loop
         
     ' row A
    k = 1
     ' iterate through the columns A, E, I that have the V-Grade Data Respectively
    Do While k < 10
        i = 1 ' for inner loop
        currentRow = 4 ' data starts at row 2
        ' t_data only iterates for data w/o blanks, currentRow keeps iterating with data and blanks
        t_data = currentRow
        ' set at first v-grade
        Set currentGradePos = ws.Cells(currentRow, k)
        ' loop through rows until 7 arrays are added to the collection
         Do While i < 8
            ' check if column k has a "V" grade
            If ws.Cells(currentRow, k).value Like "*V*" Then
                ' skip first v value
                If currentRow > 4 Then
                    
                    ' if # of blanks less than 3
                    If (currentRow - t_data) < 4 Then
                        ' insert range of blank cell so there is a total of 3 before next v-grade and increment current row back to new start
                        If (currentRow - t_data) = 1 Then
                            Set insertRange = Range(ws.Cells(currentRow, k), ws.Cells(currentRow + 2, k + numCol - 1))
                            currentRow = currentRow + 3
                        ElseIf (currentRow - t_data) = 2 Then
                            Set insertRange = Range(ws.Cells(currentRow, k), ws.Cells(currentRow + 1, k + numCol - 1))
                            currentRow = currentRow + 2
                        ElseIf (currentRow - t_data) = 3 Then
                            Set insertRange = Range(ws.Cells(currentRow, k), ws.Cells(currentRow, k + numCol - 1))
                            currentRow = currentRow + 1
                
                        End If
                        
                        ' insert specified range
                        insertRange.Insert Shift:=xlShiftDown
                        
                    ' # of blanks is greater than 3
                    ElseIf (currentRow - t_data) > 4 Then
                        ' delete all blanks in v- grade column
                        Range(ws.Cells(t_data + 1, k), ws.Cells(currentRow - 1, k + numCol - 1)).Delete Shift:=xlUp
                        
                        'set the insert range as correct value to insert 3 values
                        Set insertRange = Range(ws.Cells(t_data + 1, k), ws.Cells(t_data + 3, k + numCol - 1))
                        
                        'insert 3 blank cells
                        insertRange.Insert Shift:=xlShiftDown
                        
                        ' set currentrow as correct value after deletion and insertion
                        currentRow = t_data + 4
                                            
                    End If
                    
                    ' add border to correct range
                    Set borderRange = Range(ws.Cells(currentGradePos.Row, k), ws.Cells(currentRow - 1, k + numCol - 1))
                    AddBorder borderRange
                
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
        Loop
    
    k = k + 4 ' k will be 1, 5, 9 respectively
    
    Loop
    
End Function
Function AddBorder(selectedRange As Range)

    ' Set the outside borders (left, right, top, bottom)
    With selectedRange.borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With selectedRange.borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With selectedRange.borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With selectedRange.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Function
Sub full_Reset()
    'in case data needs to be reset set worksheet
    Set output = ActiveWorkbook.Sheets("Sort")
    startRow = 5
    
    ' reset publics
    var_Sorted = Empty
    var_Then_Sorted = Empty
    custom_Sorted = Empty
    Set loc_var_sorted = Nothing
    Set data_ws = Nothing
    ' turn off eventsfor vba on "Sort" sheet
    Application.EnableEvents = False
    output.Rows(startRow & ":" & output.Rows.Count).ClearContents ' clear only rows below row 3
    
    Cells(2, 1) = "Send Data"
    Cells(2, 3) = "---"
    Cells(2, 4) = "---"
    Cells(2, 5) = "Ascending"
    Cells(2, 6) = "---"
    Cells(2, 12) = "---"
     ' reset cell F3
     Range("F3").Validation.Delete
     Range("F3").value = ""
     ' reset cell F4
     Range("F4").Validation.Delete
     Range("F4").value = ""
     
     ' reset cell M3
     Range("M3").Validation.Delete
     Range("M3").value = ""
     ' reset cell M4
     Range("M4").Validation.Delete
     Range("M4").value = ""
     
     ' turn back on eventsfor vba on "Sort" sheet
     Application.EnableEvents = True
     
     ' Delete existing PivotCharts
    Dim pc As ChartObject
    For Each pc In output.ChartObjects
        pc.Delete
    Next pc
     'clear borders
     Call borders
     MsgBox ("Dropdowns and Data Reset") 'alert that you succesfully reset the dropdowns
End Sub
Sub borders()
     ' clear border add border
    Dim lastRow As Long
    
    If Not IsEmpty(var_Sorted) Then
        lastRow = UBound(var_Sorted) + 30 ' full length of list + 30
    Else
        lastRow = 1000 ' abritrarely large # if var_sorted is not existent
    End If
    
    ' Clear borders in the range A5 to D[lastRow]
    With Range("A5:D" & lastRow).borders
        .LineStyle = xlNone ' Removes all borders
    End With
    
    Dim firstEmptyRow As Long
    Dim dataRange As Range
    firstEmptyRow = Range("A" & startRow - 1 & ":A" & Rows.Count).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
    If firstEmptyRow = 5 Then
        '
    Else
        Set dataRange = Range("A" & startRow & ":D" & firstEmptyRow - 1)
        With dataRange.borders
            .LineStyle = xlContinuous  ' Solid line
            .Color = RGB(0, 0, 0)     ' Black color
            .Weight = xlThin          ' Thin border
        End With
    End If
End Sub

Function Dropdown(Optional ByVal sort_By As Boolean = False, _
                      Optional then_By As Boolean = False, _
                      Optional Filter As Boolean = False, _
                      Optional cust_Filter As Boolean = False, _
                      Optional asc_Desc As Boolean = False, _
                      Optional Bord As Boolean = False, _
                      Optional Stack As Boolean = False)
    ' make sure public is established
    Set output = ThisWorkbook.Worksheets("Sort")
    
    
    ' check "Sort By" dropdown
    If sort_By = True Then
        'if "Reset" is selected reset dropdowns and exit function
        If output.Cells(2, 2).value = "Reset" Then
            'sub to reset dropdowns
            Call full_Reset
            Exit Function
        ' sort by v_grade
        ElseIf output.Cells(2, 2).value = "V-Grade" Then
            Call sort_By_Rows
        ' sort by date
        ElseIf output.Cells(2, 2).value = "Date" Then
            Call mergeSort_By_Date
        ' sort by location (default is then_by date)
        ElseIf output.Cells(2, 2).value = "Location" Then
            Call loc_Sort_By_Date_Ouput
        Else
            MsgBox ("Unknown Value for *Sort By* Dropdown")
            MsgBox ("Sort Cancelled")
            Exit Function
        End If
        
    End If
    
    ' check "Then By" dropdown
    If then_By = True Then
        'do nothing if "---"
        If output.Cells(2, 3).value = "---" Then
            '
        ' sort then_By v-grade
        ElseIf output.Cells(2, 3).value = "V-Grade" Then
            then_VGrade_Sort
        ' do nothing.. Date is default then_By sort
         ElseIf output.Cells(2, 3).value = "Date" Then
            '
        ' location then_by sort not implemented yet
        ElseIf output.Cells(2, 3).value = "Location" Then
            MsgBox ("need to make a new sorting algorithm for then_by Location")
            MsgBox ("Sort Cancelled")
            Exit Function
        ' if dropdown is not recognized display message and exit function
        Else
            MsgBox ("Unknown Value for *Then By* Dropdown")
            MsgBox ("Sort Cancelled")
            Exit Function
        End If
        
    End If
    
    ' check "Filter" Dropdown
    If Filter = True Then
    
        ' if turned-off reset the stored lsit
        If output.Cells(2, 4).value = "---" Then
            emoji_Sorted = Empty
            
        ' onsight selected; Sort active list by the eagle
        ElseIf output.Cells(2, 4).value = "Onsight" Then
            Call OS_F_OneS_Sort(output.Cells(1, 8))
            
         ' flash selected; Sort active list by the flash
        ElseIf output.Cells(2, 4).value = "Flash" Then
            Call OS_F_OneS_Sort(output.Cells(1, 9))
            
        ' one sesh selected; Sort active list by the army helmet
        ElseIf output.Cells(2, 4).value = "One Sesh" Then
            Call OS_F_OneS_Sort(output.Cells(1, 10))
            
        ' project selected; Sort by no emoji
        ElseIf output.Cells(2, 4).value = "Project" Then
            Call Project_Sort
         ' if dropdown is not recognized display message and exit function
        Else
            MsgBox ("Unknown Value for *Filter* Dropdown")
            MsgBox ("Sort Cancelled")
            Exit Function
        End If
    End If
        
    ' Check Custom Filter
    If cust_Filter = True Then
    
        ' if both dropdowns have valid v-grades then filter
        If output.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
        And ValidGrade(Range("F4").value) Then
            Call custom_Grade(Range("F3").value, Range("F4").value)
            
        ' if both dropdowns have valid dates then filter
        ElseIf output.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            Call custom_Date_sort(Range("F3").value, Range("F4").value)
        
        ' if dropdown has any location then filter
        ElseIf output.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        
        End If
        
    End If
    
    ' Check Stack Filter
    If Stack = True Then
        ' if both dropdowns have valid v-grades then filter
         If output.Cells(2, 12).value = "Custom V-Grade" And ValidGrade(Range("M3").value) _
         And ValidGrade(Range("M4").value) Then
            
            Call custom_Grade(Range("M3").value, Range("M4").value, custom_Sorted)
            
        ' if both dropdowns have valid dates then filter
        ElseIf output.Cells(2, 12).value = "Custom Date" And validDate(Range("M3").value) _
        And validDate(Range("M4").value) Then
            
            Call custom_Date_sort(Range("M3").value, Range("M4").value, custom_Sorted)
         
         ' if we are sorting by location check if the cell has a value selected from the dropdown
        ElseIf output.Cells(2, 12).value = "Custom Location" And Range("M3").value <> "" Then
            Call custom_Location_Sort(Range("M3").value, custom_Sorted)
            
        End If
    End If
        
    ' check ascend / descend dropdown
    If asc_Desc = True Then
        If output.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
    End If
    
    If Bord = True Then
        Call borders
    End If
    
End Function
