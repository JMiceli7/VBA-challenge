Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

' using for each worksheet and ws before range and cells to apply across sheets

' declaring variables and adding, fitting column headers

    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Total_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim Total_Stock_Volume As Double
    Dim sheetName As String
    

    For Each ws In Worksheets

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        sheetName = ws.Name
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit

        Quarterly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
   
        Opening_Price = ws.Cells(2, 3).Value
   
   
   ' for loop to pull values for Ticker, Quarterly Change, Percent Change, and Total Stock Volume
   
                For i = 2 To lastrow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        Closing_Price = ws.Cells(i, 6).Value
                        Percent_Change = ((Closing_Price - Opening_Price) / Opening_Price)
                        Quarterly_Change = (Closing_Price - Opening_Price)
                        
                        Ticker = ws.Cells(i, 1).Value
            
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
                        ws.Range("I" & Summary_Table_Row).Value = Ticker
                        ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                        ws.Range("J" & Summary_Table_Row).NumberFormat = "$0.00"
                        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        Summary_Table_Row = Summary_Table_Row + 1
                        
                        Total_Stock_Volume = 0
                        
                        Opening_Price = ws.Cells(1 + i, 3).Value
                        
                    Else
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                    End If
                
                
                 ' color formating for Quarterly Change
                
                    If ws.Cells(Summary_Table_Row, 10) > 0 Then
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    
                    ElseIf ws.Cells(Summary_Table_Row, 10) = 0 Then ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 2
                    
                    Else
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                        
                    End If
                         
                    
                 ' color formating for Percent Change
                
                    If ws.Cells(Summary_Table_Row, 11) > 0 Then
                            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                    
                    ElseIf ws.Cells(Summary_Table_Row, 11) = 0 Then ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 2
                    
                    Else
                            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                    End If
                    
                Next i

            
    ' calculating max percent increaese, decrease, and total volume
            
    lastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Greatest_Percent_Increase = ws.Cells(2, 11).Value
    Greatest_Percent_Decrease = ws.Cells(2, 11).Value
    Greatest_Total_Volume = ws.Cells(2, 12).Value
    
   

    For j = 2 To lastTicker
        
         If ws.Cells(j, 11).Value > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = ws.Cells(j, 11).Value
         End If
        
         If ws.Cells(j, 11).Value < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = ws.Cells(j, 11).Value
         End If
        
         If ws.Cells(j, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = ws.Cells(j, 12).Value
         End If
        
    Next j



        ws.Range("Q2").Value = Greatest_Percent_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Greatest_Percent_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Greatest_Total_Volume
       
 ' pulling corresponding ticker with greatest values
      
        For k = 2 To lastTicker
        
            If ws.Cells(k, 11).Value = Greatest_Percent_Increase Then
                    ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            End If
            
            If ws.Cells(k, 11).Value = Greatest_Percent_Decrease Then
                    ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
            End If
                             
            If ws.Cells(k, 12).Value = Greatest_Total_Volume Then
                    ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
            End If
            
            
        Next k
       
    ws.Columns("I:Q").AutoFit
        
    
    Next ws
    

End Sub


Sub refresh()


    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ' Delete columns 9-17 (columns I-R)
        Range(ws.Columns(9), ws.Columns(17)).Delete
        
    Next ws
    
    
End Sub





