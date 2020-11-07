Attribute VB_Name = "Module2"
Sub yearlychange()

'Set count of rows, columns, and sheetrow
Dim lrow As Long
Dim lcol As Long

    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
Dim sheetrow As Integer
    sheetrow = 2

Dim myrange As Range

Dim minopen As Long
Dim maxclose As Long

Dim mindate As Object
Dim maxdate As Object


Dim yearlychange As Double

For i = 2 To lrow

    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
        mindate = Application.WorksheetFunction.Min(Cells(i, 2))
        
        maxdate = WorksheetFunction.Max(Cells(i, 2))
        
        minopen = WorksheetFunction.Match((Min), Cells(i, 3))
        
        maxclose = WorksheetFunction.Match((Max), Cells(i, 6))
        
        
        yearlychange = maxclose - minopen
        
        
        
        'Print
        Range("J" & sheetrow).Value = yearlychange
        
        sheetrow = sheetrow + 1
        
        'yearlychange = 0
        
        
    Else
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Print
        Range("J" & sheetrow).Value = yearlychange
        
        sheetrow = sheetrow + 1
        
        'yearlychange = 0
        
    
    End If
    
    End If
    
    Next i
    


End Sub
