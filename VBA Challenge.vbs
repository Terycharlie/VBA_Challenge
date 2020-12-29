Attribute VB_Name = "Module1"
Sub ticker()

Dim ticker As String
Dim ws As Worksheet
Dim wsCount As Integer
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Volume As Double
Dim Lastrow As Long
Dim Beg_Price As Double
Dim Summary_Table_Row As Long



    'worksheet iterate
wsCount = ActiveWorkbook.Worksheets.Count


    'Insert headers in each worksheet
For Each ws In ThisWorkbook.Worksheets

   ws.Range("i1").Value2 = "Ticker"
   ws.Range("J1").Value2 = "Yearly_Change"
   ws.Range("K1").Value2 = "Percentage_Change"
   ws.Range("L1").Value2 = "Total Volume"
   
   
   'Setup location for variables
   Summary_Table_Row = 2
    
   'Define lastrow
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    Beg_Price = 2
    
    ' Loop through all worksheets to populate data
    For i = 2 To Lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'find all the values
            ticker = ws.Cells(i, 1).Value
            Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(Beg_Price, 6).Value
            
            'To avoid getting an error when dividing by zero
            If ws.Cells(Beg_Price, 6).Value = 0 Then
            Percentage_Change = 0
        Else
                Percentage_Change = Round(Yearly_Change / ws.Cells(Beg_Price, 6).Value, 4)
            
     End If
            
           
            'Populate the Summary Table
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 12).Value = Total_Volume
            ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
            ws.Cells(Summary_Table_Row, 11).Value = Percentage_Change
            
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Volume = 0
                Beg_Price = i + 1
            
            
        Else
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                   
    
    End If
    
Next i
 ws.Range("K:K").NumberFormat = ("0.00%")
 
   'Conditional Formatting_Percentage Change
    Dim PerChange As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set PerChange = ws.Range("J2", ws.Range("J2").End(xlDown))
    c = PerChange.Cells.Count
    
    For g = 1 To c
        Set color_cell = PerChange(g)
        Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g

 
Next ws

End Sub

