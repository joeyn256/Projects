VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Worksheet_Change(ByVal Target As Range)
    
    startRow = 5 ' Incase this is the first function called, we declare our publics
    numCol = 4
    ' logic and arguments for custom sort with v-grade
    Dim arg1 As String
    Dim arg2 As String
    
    ' if dropdown list is changed for pulling from Send Data Or Project Data Spreadsheet
    If Not Intersect(Target, Range("A2")) Is Nothing Then
        
        ' check "Sort By" dropdown
        If Me.Cells(2, 2).value = "V-Grade" Then
            Call sort_By_Rows
        ElseIf Me.Cells(2, 2).value = "Date" Then
            Call mergeSort_By_Date
        ElseIf Me.Cells(2, 2).value = "Location" Then
            Call loc_Sort_By_Date_Ouput
        End If
        
        ' check "Then By" dropdown
        If Me.Cells(2, 3).value = "V-Grade" Then
            then_VGrade_Sort
        End If
            
        ' check "Filter" Dropdown
        If Me.Cells(2, 4).value = "---" Then
            '
        ElseIf Me.Cells(2, 4).value = "Onsight" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 8))
        ElseIf Me.Cells(2, 4).value = "Flash" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 9))
        ElseIf Me.Cells(2, 4).value = "One Sesh" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 10))
        ElseIf Me.Cells(2, 4).value = "Project" Then
            Call Project_Sort
        End If
        
        ' Check Custom Filter
        If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
        And ValidGrade(Range("F4").value) Then
            Call custom_Grade(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            Call custom_Date_sort(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders
        
    End If
    
    ' if dropdown list is changed for Sort_By
    If Not Intersect(Target, Range("B2")) Is Nothing Then
        Select Case Target.value

        ' check if the changed cell is the drop-down cell
        Case "Reset"
            'sub to reset dropdowns
            Call full_Reset
            Exit Sub
        
        ' sort by v_grade being aware of what the other values are
        Case "V-Grade"
            ' no matter what we refresh the list
            Call sort_By_Rows
            
        Case "Date"
            ' no matter what when this is selected we refresh sort_By_Date
            Call mergeSort_By_Date
                    
        Case "Location"
            Call loc_Sort_By_Date_Ouput
            
        End Select
        
        ' check "Then By" dropdown
        If Me.Cells(2, 3).value = "V-Grade" Then
            then_VGrade_Sort
        End If
        
        ' check "Filter" Dropdown
        If Me.Cells(2, 4).value = "---" Then
            '
        ElseIf Me.Cells(2, 4).value = "Onsight" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 8))
        ElseIf Me.Cells(2, 4).value = "Flash" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 9))
        ElseIf Me.Cells(2, 4).value = "One Sesh" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 10))
        ElseIf Me.Cells(2, 4).value = "Project" Then
            Call Project_Sort
        End If
        
        ' Check Custom Filter
        If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
        And ValidGrade(Range("F4").value) Then
            Call custom_Grade(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            Call custom_Date_sort(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        End If
        
         ' check ascend / descend dropdown
        If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders

    End If
    
    If Not Intersect(Target, Range("C2")) Is Nothing Then ' When dropdown list is changed for sort Then_By
        
        ' check "Sort By" dropdown
        If Me.Cells(2, 2).value = "V-Grade" Then
            Call sort_By_Rows
        ElseIf Me.Cells(2, 2).value = "Date" Then
            Call mergeSort_By_Date
        ElseIf Me.Cells(2, 2).value = "Location" Then
            Call loc_Sort_By_Date_Ouput
        End If
        
        Select Case Target.value
            Case "---"
                ' do nothing
            
            Case "V-Grade"
                ' call "Then By" for "V-Grade"
                Call then_VGrade_Sort
            
            Case "Date"
                ' do nothing this is how we sort it
                
             Case "Location"
                MsgBox ("need to make a new sorting algorithm for then_by Location")
                
                Exit Sub

            Case Else
                ' If none of the cases match, do nothing or add another behavior
                Exit Sub
        End Select
        
         ' check "Filter" Dropdown
        If Me.Cells(2, 4).value = "---" Then
            '
        ElseIf Me.Cells(2, 4).value = "Onsight" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 8))
        ElseIf Me.Cells(2, 4).value = "Flash" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 9))
        ElseIf Me.Cells(2, 4).value = "One Sesh" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 10))
        ElseIf Me.Cells(2, 4).value = "Project" Then
            Call Project_Sort
        End If
        
        ' Check Custom Filter
        If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
        And ValidGrade(Range("F4").value) Then
            Call custom_Grade(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            Call custom_Date_sort(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
    
        'add borders
        Call borders
            
    End If
    
    If Not Intersect(Target, Range("D2")) Is Nothing Then
        Select Case Target.value
        'reset emoji sort
        Case "---"
            'reset the public variable emoji sort so it doesn't hold the previous values
            emoji_Sorted = Empty
            
            ' check "Sort By" dropdown
            If Me.Cells(2, 2).value = "V-Grade" Then
                Call sort_By_Rows
            ElseIf Me.Cells(2, 2).value = "Date" Then
                Call mergeSort_By_Date
            ElseIf Me.Cells(2, 2).value = "Location" Then
                Call loc_Sort_By_Date_Ouput
            End If
            
             ' check "Then By" dropdown
            If Me.Cells(2, 3).value = "V-Grade" Then
                then_VGrade_Sort
            End If
        
        Case "Onsight"
            ' Sort active list by the eagle
            Call OS_F_OneS_Sort(Me.Cells(1, 8))
                
        Case "Flash"
            ' Sort active list by the flash
            Call OS_F_OneS_Sort(Me.Cells(1, 9))
            
        Case "One Sesh"
            ' Sort active list by the helmet
            Call OS_F_OneS_Sort(Me.Cells(1, 10))
        
            
        Case "Project"
            Call Project_Sort

        End Select
        
        ' Check Custom Filter
        If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
        And ValidGrade(Range("F4").value) Then
            Call custom_Grade(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            Call custom_Date_sort(Range("F3").value, Range("F4").value)
        ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders
        
    End If
    
    
    If Not Intersect(Target, Range("E2")) Is Nothing Then
        Select Case Target.value
      
        Case "Ascending"
            Call ascend
        
        Case "Descending"
            Call descend
            
        End Select
        
    End If
    
    If Not Intersect(Target, Range("F2")) Is Nothing Then
         ' check "Sort By" dropdown
            If Me.Cells(2, 2).value = "V-Grade" Then
                Call sort_By_Rows
            ElseIf Me.Cells(2, 2).value = "Date" Then
                Call mergeSort_By_Date
            ElseIf Me.Cells(2, 2).value = "Location" Then
                Call loc_Sort_By_Date_Ouput
            End If
            
            ' check "Then By" dropdown
            If Me.Cells(2, 3).value = "V-Grade" Then
                then_VGrade_Sort
            End If
                    
            ' check "Filter" Dropdown
            If Me.Cells(2, 4).value = "---" Then
                '
            ElseIf Me.Cells(2, 4).value = "Onsight" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 8))
            ElseIf Me.Cells(2, 4).value = "Flash" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 9))
            ElseIf Me.Cells(2, 4).value = "One Sesh" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 10))
            ElseIf Me.Cells(2, 4).value = "Project" Then
                Call Project_Sort
            End If
        
        ' make the correct dropdowns
        Select Case Target.value
        Case "---"
            ' reset cell F3
             Range("F3").Validation.Delete
             Range("F3").value = ""
             ' reset cell F4
             Range("F4").Validation.Delete
             Range("F4").value = ""
               
        Case "Custom V-Grade"
            'add appropriate dropdowns
            Call addDataValidation("F3", "F4")
            
        Case "Custom Date"
            Call addDataValidation_Dates("F3", "F4")
            
        Case "Custom Location"
            Call addDataValidation_Location("F3", "F4")
        End Select
        
        'make sure sorted list is correct
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
    
        'add borders
        Call borders
    End If
    
     If Not Intersect(Target, Range("F3")) Is Nothing Or Not Intersect(Target, Range("F4")) Is Nothing Then
        ' if both dropdowns have valid v-grades then filter
         If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
         And ValidGrade(Range("F4").value) Then
            
            Call custom_Grade(Range("F3").value, Range("F4").value)
            
        ' if both dropdowns have valid dates then filter
        ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
            
            Call custom_Date_sort(Range("F3").value, Range("F4").value)

        ' if we are sorting by location check if the cell has a value selected from the dropdown
        ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
            Call custom_Location_Sort(Range("F3").value)
        
        End If
        
        
        ' check is sorted by ascending or descending
        If Me.Cells(2, 5).value = "Descending" Then
            Call descend
        End If
     
     End If
     
     ' dropdown for stack sort
     If Not Intersect(Target, Range("L2")) Is Nothing Then
        ' check previous sorts
        ' check "Sort By" dropdown
            If Me.Cells(2, 2).value = "V-Grade" Then
                Call sort_By_Rows
            ElseIf Me.Cells(2, 2).value = "Date" Then
                Call mergeSort_By_Date
            ElseIf Me.Cells(2, 2).value = "Location" Then
                Call loc_Sort_By_Date_Ouput
            End If
            
            ' check "Then By" dropdown
            If Me.Cells(2, 3).value = "V-Grade" Then
                then_VGrade_Sort
            End If
                    
            ' check "Filter" Dropdown
            If Me.Cells(2, 4).value = "---" Then
                '
            ElseIf Me.Cells(2, 4).value = "Onsight" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 8))
            ElseIf Me.Cells(2, 4).value = "Flash" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 9))
            ElseIf Me.Cells(2, 4).value = "One Sesh" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 10))
            ElseIf Me.Cells(2, 4).value = "Project" Then
                Call Project_Sort
            End If
            
            ' Check Custom Filter
            If Me.Cells(2, 6).value = "Custom V-Grade" And ValidGrade(Range("F3").value) _
            And ValidGrade(Range("F4").value) Then
                Call custom_Grade(Range("F3").value, Range("F4").value)
            ElseIf Me.Cells(2, 6).value = "Custom Date" And validDate(Range("F3").value) And validDate(Range("F4").value) Then
                Call custom_Date_sort(Range("F3").value, Range("F4").value)
            ElseIf Me.Cells(2, 6).value = "Custom Location" And Range("F3").value <> "" Then
                Call custom_Location_Sort(Range("F3").value)
            End If
            
            'call descend before so we don't call it with custom_sorted empty
            If Me.Cells(2, 5).value = "Descending" And Not IsEmpty(var_Sorted) Then
                Call descend
            End If
            'add borders
            Call borders
            
            ' make sure correct dropdowns are added
            Select Case Target.value
            Case "---"
                ' reset cell F3
                 Range("M3").Validation.Delete
                 Range("M3").value = ""
                 ' reset cell F4
                 Range("M4").Validation.Delete
                 Range("M4").value = ""
                 custom_Sorted = Empty
                
            ' add appropriate dropdowns
            Case "Custom V-Grade"
                Call addDataValidation("M3", "M4")
                
            Case "Custom Date"
                Call addDataValidation_Dates("M3", "M4")
                
            Case "Custom Location"
                Call addDataValidation_Location("M3", "M4", custom_Sorted)
            End Select
            ' resume checking correct sort
        End If
    
    If Not Intersect(Target, Range("M3")) Is Nothing Or Not Intersect(Target, Range("M4")) Is Nothing Then
        ' if both dropdowns have valid v-grades then filter
         If Me.Cells(2, 12).value = "Custom V-Grade" And ValidGrade(Range("M3").value) _
         And ValidGrade(Range("M4").value) Then
            
            Call custom_Grade(Range("M3").value, Range("M4").value, custom_Sorted)
            
        ' if both dropdowns have valid dates then filter
        ElseIf Me.Cells(2, 12).value = "Custom Date" And validDate(Range("M3").value) And validDate(Range("M4").value) Then
            
            Call custom_Date_sort(Range("M3").value, Range("M4").value, custom_Sorted)
         
         ' if we are sorting by location check if the cell has a value selected from the dropdown
        ElseIf Me.Cells(2, 12).value = "Custom Location" And Range("M3").value <> "" Then
            Call custom_Location_Sort(Range("M3").value, custom_Sorted)
            
        End If
        
        If Me.Cells(2, 5).value = "Descending" Then
                Call descend
            End If
     
     End If
        
End Sub

