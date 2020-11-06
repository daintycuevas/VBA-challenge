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

Next ws

MsgBox ("Okay")


End Sub




