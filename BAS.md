Attribute VB_Name = "Module1"
Sub Module_2_Homework()
    
    For Each ws In Worksheets
    ws.Activate
    
    'Delcare Variables
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim LastRow As Integer
    Dim Total_Stock_Volume As LongLong
    Dim Max_Percent As String
    Dim Min_Percent As String
    Dim Greatest_Value As LongLong
 
    
    'Loop to the bottom of the data column and save as LastRow
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Opening_Value = Cells(2, 3).Value
    
    'Loop through all rows in column A
    For i = 2 To LastRow
    
        'If the ticker is different than the previous row then
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            'Set the variable names and perform algebra
            Ticker_Name = Cells(i, 1).Value
            Closing_Value = Cells(i, 6).Value
            Yearly_Change = (Closing_Value - Opening_Value)
            Percent_Change = (Yearly_Change / Opening_Value)
            Opening_Value = Cells(i + 1, 3).Value
            Total_Stock_Volume = Total_Stock_Value + Cells(i, 7).Value
            
            'Assign columns to Summary_Table_Row
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            'Reset Total_Stock_Volume after each row
            Total_Stock_Volume = 0
            
            'Format percentage
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
        
    Else
    Total_Stock_Volume = Total_Stock_Value + Cells(i, 7).Value
    
    End If
    
    Next i
    
      'Insert max percent increase into Range "P2"
      Max_Percent = WorksheetFunction.Max(Range("K:K"))
      Cells(2, 16).Value = Max_Percent
      Cells(2, 16).NumberFormat = "0.00%"
    
      increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
      Cells(2, 15).Value = Cells(increase_number + 1, 9)
      
      'Insert max percent decrease into range "P3"
       Min_Percent = WorksheetFunction.Min(Range("K:K"))
       Cells(3, 16).Value = Min_Percent
       Cells(3, 16).NumberFormat = "0.00%"
       
      decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
      Cells(3, 15).Value = Cells(decrease_number + 1, 9)
       
       'Insert greatest total volume
       Greatest_Value = WorksheetFunction.Max(Range("L:L"))
       Cells(4, 16).Value = Greatest_Value
       
       TotalVol_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
       Cells(4, 15).Value = Cells(TotalVol_number + 1, 9)
       
    'Add Color
     For i = 2 To 122
     
     If Cells(i, 11).Value < 0 Then
        Cells(i, 11).Interior.ColorIndex = 3
        
    Else: Cells(i, 11).Interior.ColorIndex = 4
    
    End If
      
    Next i
    
    'Add Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Next ws
    
    
    End Sub