Attribute VB_Name = "Module3"
Sub color()

'Set count of rows, columns, and sheetrow
Dim lrow As Long
Dim lcol As Long

    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
Dim sheetrow As Integer
    sheetrow = 2

For i = 2 To lrow

    If Range("K") < 0 Then
        Range("K").Interior.ColorIndex = 3
        
    End If
    
    If Range("K") > 0 Then
        Range("K").Interior.ColorIndex = 4
        
    End If
    
    Next i
    

End Sub
