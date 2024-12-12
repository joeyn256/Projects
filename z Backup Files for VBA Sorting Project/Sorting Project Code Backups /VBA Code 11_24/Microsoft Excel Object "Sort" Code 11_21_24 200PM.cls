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
        If Me.Cells(2, 2).Value = "V-Grade" Then
            Call sort_By_Rows
        ElseIf Me.Cells(2, 2).Value = "Date" Then
            Call mergeSort_By_Date
        ElseIf Me.Cells(2, 2).Value = "Location" Then
            Call loc_Sort_By_Date_Ouput
        End If
        
        ' check "Then By" dropdown
        If Me.Cells(2, 3).Value = "---" Then
            '
        ElseIf Me.Cells(2, 3).Value = "V-Grade" Then
            then_VGrade_Sort
        End If
            
        ' check "Filter" Dropdown
        If Me.Cells(2, 4).Value = "---" Then
            '
        ElseIf Me.Cells(2, 4).Value = "Onsight" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 8))
        ElseIf Me.Cells(2, 4).Value = "Flash" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 9))
        ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
            Call OS_F_OneS_Sort(Me.Cells(1, 10))
        ElseIf Me.Cells(2, 4).Value = "Project" Then
            Call Project_Sort
        End If
        
        ' check custom filter
        If Me.Cells(2, 6).Value = "Custom V-Grade" And ValidGrade(Range("F3").Value) _
        And ValidGrade(Range("F4").Value) Then
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).Value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders
        
    End If
    
    ' if dropdown list is changed for Sort_By
    If Not Intersect(Target, Range("B2")) Is Nothing Then
        Select Case Target.Value

        ' check if the changed cell is the drop-down cell
        Case "Clear"
            'ws1.cells.clear clear any previous data on sheet, "Sort"
            Me.Rows(startRow & ":" & Me.Rows.Count).ClearContents ' clear only rows below row 3
            ' reset publics
            Call borders
            var_Sorted = Empty
            var_Then_Sorted = Empty
            Set loc_var_sorted = Nothing
            Set data_ws = Nothing
             ' reset cell F3
             Range("F3").Validation.Delete
             Range("F3").Value = ""
             ' reset cell F4
             Range("F4").Validation.Delete
             Range("F4").Value = ""
            'clear borders
            MsgBox ("Successfully Cleared") 'alert you that the list is cleared
            
            Exit Sub
        
        ' sort by v_grade being aware of what the other values are
        Case "V-Grade"
            ' no matter what we refresh the list
            Call sort_By_Rows
            
            ' check "Then By" dropdown
            If Me.Cells(2, 3).Value = "---" Then
                '
            ElseIf Me.Cells(2, 3).Value = "V-Grade" Then
                then_VGrade_Sort
            End If
            
            ' check "Filter" Dropdown
            If Me.Cells(2, 4).Value = "---" Then
                '
            ElseIf Me.Cells(2, 4).Value = "Onsight" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 8))
            ElseIf Me.Cells(2, 4).Value = "Flash" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 9))
            ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 10))
            ElseIf Me.Cells(2, 4).Value = "Project" Then
                Call Project_Sort
            End If
            
        Case "Date"
            ' no matter what when this is selected we refresh sort_By_Date
            Call mergeSort_By_Date
            
            ' check "Then By" dropdown
            If Me.Cells(2, 3).Value = "---" Then
                '
            ElseIf Me.Cells(2, 3).Value = "V-Grade" Then
                then_VGrade_Sort
            End If
            
            ' check "Filter" Dropdown
            If Me.Cells(2, 4).Value = "---" Then
                '
           ElseIf Me.Cells(2, 4).Value = "Onsight" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 8))
            ElseIf Me.Cells(2, 4).Value = "Flash" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 9))
            ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 10))
            ElseIf Me.Cells(2, 4).Value = "Project" Then
                Call Project_Sort
            End If
            
        Case "Location"
            Call loc_Sort_By_Date_Ouput
            
            ' check "Then By" dropdown
            If Me.Cells(2, 3).Value = "---" Then
                '
            ElseIf Me.Cells(2, 3).Value = "V-Grade" Then
                then_VGrade_Sort
            End If
            
            ' check "Filter" Dropdown
            If Me.Cells(2, 4).Value = "---" Then
                '
            ElseIf Me.Cells(2, 4).Value = "Onsight" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 8))
            ElseIf Me.Cells(2, 4).Value = "Flash" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 9))
            ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
                Call OS_F_OneS_Sort(Me.Cells(1, 10))
            ElseIf Me.Cells(2, 4).Value = "Project" Then
                Call Project_Sort
            End If

        End Select
        
        ' check custom filter
        If Me.Cells(2, 6).Value = "Custom V-Grade" And ValidGrade(Range("F3").Value) _
        And ValidGrade(Range("F4").Value) Then
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
        End If
        
         ' check ascend / descend dropdown
        If Me.Cells(2, 5).Value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders

    End If
    
    If Not Intersect(Target, Range("C2")) Is Nothing Then ' When dropdown list is changed for sort Then_By
        Select Case Target.Value
            Case "---"
                ' check "Sort By" dropdown
                If Me.Cells(2, 2).Value = "V-Grade" Then
                    Call sort_By_Rows
                ElseIf Me.Cells(2, 2).Value = "Date" Then
                    Call mergeSort_By_Date
                ElseIf Me.Cells(2, 2).Value = "Location" Then
                    Call loc_Sort_By_Date_Ouput
                End If
                
                ' check "Filter" Dropdown
                If Me.Cells(2, 4).Value = "---" Then
                    '
                ElseIf Me.Cells(2, 4).Value = "Onsight" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 8))
                ElseIf Me.Cells(2, 4).Value = "Flash" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 9))
                ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 10))
                ElseIf Me.Cells(2, 4).Value = "Project" Then
                    Call Project_Sort
                End If
            
            Case "V-Grade"
                
                ' check "Sort By" dropdown
                If Me.Cells(2, 2).Value = "V-Grade" Then
                    Call sort_By_Rows
                ElseIf Me.Cells(2, 2).Value = "Date" Then
                    Call mergeSort_By_Date
                ElseIf Me.Cells(2, 2).Value = "Location" Then
                    Call loc_Sort_By_Date_Ouput
                End If
                
                ' call "Then By" for "V-Grade"
                Call then_VGrade_Sort
                
                ' check "Filter" Dropdown
                If Me.Cells(2, 4).Value = "---" Then
                    '
                ElseIf Me.Cells(2, 4).Value = "Onsight" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 8))
                ElseIf Me.Cells(2, 4).Value = "Flash" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 9))
                ElseIf Me.Cells(2, 4).Value = "One Sesh" Then
                    Call OS_F_OneS_Sort(Me.Cells(1, 10))
                ElseIf Me.Cells(2, 4).Value = "Project" Then
                    Call Project_Sort
                End If
            
            Case "Date (to be implemented)"
                MsgBox ("need to make a new sorting algorithm for then_by Date")
                
             Case "Location (to be implemented)"
                MsgBox ("need to make a new sorting algorithm for then_by Location")

            Case Else
                ' If none of the cases match, do nothing or add another behavior
                Exit Sub
        End Select
        
        ' check custom filter
        If Me.Cells(2, 6).Value = "Custom V-Grade" And ValidGrade(Range("F3").Value) _
        And ValidGrade(Range("F4").Value) Then
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).Value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
    
        'add borders
        Call borders
            
    End If
    
    If Not Intersect(Target, Range("D2")) Is Nothing Then
        Select Case Target.Value
        'reset emoji sort
        Case "---"
            'reset the public variable emoji sort so it doesn't hold the previous values
            emoji_Sorted = Empty
            
            ' check "Sort By" dropdown
            If Me.Cells(2, 2).Value = "V-Grade" Then
                Call sort_By_Rows
            ElseIf Me.Cells(2, 2).Value = "Date" Then
                Call mergeSort_By_Date
            ElseIf Me.Cells(2, 2).Value = "Location" Then
                Call loc_Sort_By_Date_Ouput
            End If
            
             ' check "Then By" dropdown
            If Me.Cells(2, 3).Value = "---" Then
                '
            ElseIf Me.Cells(2, 3).Value = "V-Grade" Then
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
        
        ' check custom filter
        If Me.Cells(2, 6).Value = "Custom V-Grade" And ValidGrade(Range("F3").Value) _
        And ValidGrade(Range("F4").Value) Then
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
        End If
        
        ' check ascend / descend dropdown
        If Me.Cells(2, 5).Value = "Descending" And Not IsEmpty(var_Sorted) Then
            Call descend
        End If
        
        'add borders
        Call borders
        
    End If
    
    
    If Not Intersect(Target, Range("E2")) Is Nothing Then
        Select Case Target.Value
      
        Case "Ascending"
            Call ascend
        
        Case "Descending"
            Call descend
            
        End Select
        
    End If
    
    If Not Intersect(Target, Range("F2")) Is Nothing Then
        Select Case Target.Value
        Case "---"
            ' reset cell F3
             Range("F3").Validation.Delete
             Range("F3").Value = ""
             ' reset cell F4
             Range("F4").Validation.Delete
             Range("F4").Value = ""
        Case "Custom V-Grade"
            'add appropriate dropdowns
            Call addDataValidation
            
        Case "Custom Date"
            MsgBox ("not implemented yet")
        Case "Custom Location"
            MsgBox ("not implemented yet")
        End Select
    End If
    
    If Not Intersect(Target, Range("F3")) Is Nothing Then
        ' if both dropdowns have valid v-grades then filter
        If ValidGrade(Range("F3").Value) And ValidGrade(Range("F4").Value) Then
            
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
            
             ' check is sorted by ascending or descending
            If Me.Cells(2, 5).Value = "Descending" Then
                Call descend
            End If
            
            'add borders
            Call borders
            
        End If
        
    End If
    
     If Not Intersect(Target, Range("F4")) Is Nothing Then
        ' if both dropdowns have valid v-grades then filter
         If ValidGrade(Range("F3").Value) And ValidGrade(Range("F4").Value) Then
            
            Call custom_Grade(Range("F3").Value, Range("F4").Value)
            
            ' check ascend / descend dropdown
            If Me.Cells(2, 5).Value = "Descending" Then
                Call descend
            End If
            
            'add borders
            Call borders
            
        End If
     
     End If
        
        
        
End Sub

