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
    
    ' define the dropdown ranges (A2, B2, C2, D2)
    Dim DropdownRanges As Range
    Set DropdownRanges = Union(Range("A2"), Range("B2"), Range("C2"), Range("D2"))
    
    ' if (A2, B2, C2, D2) dropdowns are changed then check the relevant tabs
    If Not Intersect(Target, DropdownRanges) Is Nothing Then
       ' call "Sort By", "Then By", "Filter", "Custom Filter", "Asc/Desc", "Borders"
        Call Dropdown(True, True, True, True, True, True)
    End If
    
    ' When dropdown list is changed for Ascending / Descending
    If Not Intersect(Target, Range("E2")) Is Nothing Then
        'call "Asc/Desc"
        Call Dropdown(, , , , True)
    End If
    
    'custom filter is changed, reset and add correct Dropdowns
    If Not Intersect(Target, Range("F2")) Is Nothing Then
        ' Re-sort without custom filter
        ' call "Sort By", "Then By", "Filter", "Asc/Desc", "Borders"
        Call Dropdown(True, True, True, , True, True)
        
        ' Insert correct Dropdown Data Validation
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
       
    End If
    
    If Not Intersect(Target, Range("F3")) Is Nothing Or Not Intersect(Target, Range("F4")) Is Nothing Then
          ' call "Custom Filter", "Asc/Desc", "Borders"
         Call Dropdown(, , , True, True, True)
     
     End If
     
     ' dropdown for stack sort
     If Not Intersect(Target, Range("L2")) Is Nothing Then
        Select Case Target.value
            ' if "---" selected then RESET sort
            Case "---"
                ' reset cell F3
                 Range("M3").Validation.Delete
                 Range("M3").value = ""
                 ' reset cell F4
                 Range("M4").Validation.Delete
                 Range("M4").value = ""
                
                ' call "Sort By", "Then By", "Filter", "Custom Filter", "Asc/Desc", "Borders"
                Call Dropdown(True, True, True, True, True, True)
                
            ' add appropriate dropdowns
            Case "Custom V-Grade"
                Call addDataValidation("M3", "M4")
                
            Case "Custom Date"
                Call addDataValidation_Dates("M3", "M4")
                
            Case "Custom Location"
                Call addDataValidation_Location("M3", "M4", custom_Sorted)
        
        End Select
            
    End If
    
    If Not Intersect(Target, Range("M3")) Is Nothing Or Not Intersect(Target, Range("M4")) Is Nothing Then
        
        ' call "Stack Sort"
        Call Dropdown(, , , , , , True)
        ' call "Asc/Desc", "Borders"
        Call Dropdown(, , , , True, True)
     
     End If
        
End Sub

