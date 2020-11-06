Attribute VB_Name = "Module1"
Sub test()

'Loop through all sheets
For Each ws In Worksheets

Dim WorksheetName As String

WorksheetName = ws.Name


'Set count of rows, columns, and sheetrow
Dim lrow As Long
Dim lcol As Long

    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
 
 Dim sheetrow As Integer
    sheetrow = 2
    
 
 'Create variables
 Dim ticker As String
 Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
 Dim yearlychange As Double
 
 Dim firstopen As String
 Dim lastclose As String
 
For i = 2 To lrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
        
        'Add TotalStockVolume
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        'Print ticker value
        ws.Range("I" & sheetrow).Value = ticker
        
        'Print the TotalStockVolume
        ws.Range("L" & sheetrow).Value = TotalStockVolume
        
        'Add one to sheetrow
        sheetrow = sheetrow + 1
        
        'Reset the TotalStockVolume
        TotalStockVolume = 0
        
    'If cell immediately following a row is the same ticker
    Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
    End If
    
Next i

For i = 2 To lrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        firstopen = Range("A2:A", FirstCell).Select
        
        lastclose = Range("F2:F", LastCell).Select
        
        yearlychange = lastclose - firstopen
        
        'Print Yearly Change value
        Range("J" & sheetrow).Value = yearlychange
        
        'Add one to sheetrow
        sheetrow = sheetrow + 1
        
        'Reset yearlychange
        yearlychange = 0
   
   End If

Next ws

MsgBox ("Okay")


End Sub




