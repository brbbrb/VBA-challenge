Attribute VB_Name = "Module1"
Sub StockData()
    
    '-------------------------
    'SETUP VARIABLES FOR CHART
    '-------------------------
        
    Dim Ticker_Symbol As String
    
    Dim Year_Change As Double
    
    Dim Percent_Change As Double
    
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    
    'Determine the first row Ticker Symbol shows up on
    Dim First_Ticker_Row As LongLong
    First_Ticker_Row = 2
    
    'Determine Last Row of table
    Dim LastRow As LongLong
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Keep track of row location for each Ticker Symbol in summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Add the Headers to Summary Table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Loop through all ticker symbols
    For i = 2 To LastRow
    
        'Check if we are still within the same Ticker Symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
           'Set the Ticker Symbol
           Ticker_Symbol = Cells(i, 1).Value
           
           'Add to Total Stock Volume
           Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
               
           'Print New Ticker Symbol in Summary Table
           Range("I" & Summary_Table_Row).Value = Ticker_Symbol
           
           'Print volume to Total Stock Volume for Ticker Symbol
           Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
           
           'Calculate & Print Yearly Change
           Range("J" & Summary_Table_Row).Value = Cells(i, 6).Value - Cells(First_Ticker_Row, 3).Value
           
           'Skip over if opening price is 0
           If Cells(First_Ticker_Row, 3).Value <> 0 Then
           
                'Calculate Percentage Change
                Range("K" & Summary_Table_Row).Value = ((Cells(i, 6).Value - Cells(First_Ticker_Row, 3).Value) / Cells(First_Ticker_Row, 3).Value)
            
           Else
           
                Range("K" & Summary_Table_Row).Value = " "
             
            End If
            
        'Add a row to Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset Total Stock Volume
        Total_Stock_Volume = 0
        
        'Reset First Ticker Row value to next row
        First_Ticker_Row = i + 1
        
        'If the cell immediately following a row is not the same Ticker Symbol
        ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
            'Add to the Total_Stock_Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
     
        End If
        
        If Cells(Summary_Table_Row, 10).Value > 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
        ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
        Else
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 6
            
        End If
        
    Next i
    
Range("K2:K" & LastRow).NumberFormat = "0.00%"

MsgBox ("DONE!")

End Sub
