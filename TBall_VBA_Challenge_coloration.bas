Attribute VB_Name = "Module2"
Sub Coloration()

    For Each ws In Worksheets
    
    'Define a variable for column to identify the yearly_change column
    Dim column As Integer
    column = 10
    
    last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'For loop thru rows in the column
    For i = 2 To last_row
    
        
        'if function to determine red cells
        If ws.Cells(i, column) < 0 Then
            ws.Cells(i, column).Interior.ColorIndex = 3
            
            
            'Elseif function to determine green cells
            ElseIf ws.Cells(i, column) >= 0 Then
                ws.Cells(i, column).Interior.ColorIndex = 4
                
                
        End If
        
    Next i
    
    Next

End Sub

