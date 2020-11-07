Attribute VB_Name = "Module2"
Sub yearlychange()

'Set count of rows, columns, and sheetrow
Dim lrow As Long
Dim lcol As Long

    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
Dim sheetrow As Integer
    sheetrow = 2

Dim firstopen As String

For i = 2 To lrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
        firstopen = Cells(i, 3).Value
  
        Range("J" & sheetrow).Value = firstopen
           
        sheetrow = sheetrow + 1
 

    End If

Next i

End Sub
