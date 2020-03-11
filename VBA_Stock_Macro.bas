Attribute VB_Name = "Module1"
Sub VBAStocks()

'Set variable to run for each worksheet
Dim ws As Worksheet

'For loop to run process for each worksheet in workbook

For Each ws In Worksheets

'Determine necessary Variables

        Dim lastrow As Long
        Dim i_summary As Integer
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim TopStock As String
        Dim TopStockChange As Double
        Dim WorstStock As String
        Dim WorstStockChange As Double
        Dim MostStockVolume As String
        Dim MostStockVolumeAmt As Double
    

'Initial Variable Values
        TotalVolume = 0
        i_summary = 2
        TopStockChange = ws.Range("K2").Value
        WorstStockChange = ws.Range("K2").Value
        MostStockVolumeAmt = ws.Range("L2").Value
    
'Output headers for output and summary table
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("M1").Value = "Opening Price"
        ws.Range("N1").Value = "Closing Price"
        ws.Range("Q1").Value = "Ticker Symbol"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest Percent Increase"
        ws.Range("P3").Value = "Greatest Percent Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
    

'LastRow Value

        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
'For Loop to output unique stock ticker symbol and total stock volume

    For i = 2 To lastrow
  
   'If Statement to determine each stock's opening price
   
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        ws.Cells(i_summary, 13).Value = ws.Cells(i, 3).Value
        End If
    
    'If statement to determine each stock's total volume
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       
    'Place the stock name on output table
    
        ws.Cells(i_summary, 9).Value = ws.Cells(i, 1).Value
    
    'Place year end closing price on the table
    
        ws.Cells(i_summary, 14).Value = ws.Cells(i, 6).Value
    
    'Place the total volume on the output table
        ws.Cells(i_summary, 12).Value = TotalVolume
    
    'Place the yearly change on the output table and add in conditional formatting based on positive/negative change. Any 0s are left uncolored
    
        ws.Cells(i_summary, 10).Value = ws.Cells(i_summary, 14).Value - ws.Cells(i_summary, 13).Value
     
        If ws.Cells(i_summary, 10).Value < 0 Then
            ws.Cells(i_summary, 10).Interior.ColorIndex = 3
        Else
            If ws.Cells(i_summary, 10).Value > 0 Then
            ws.Cells(i_summary, 10).Interior.ColorIndex = 4
            End If
        End If
     
    'Calculate percent change and Place percent change on the output table. Need to account for any closing stock price is 0 using If statement
    
        If ws.Cells(i_summary, 13).Value = 0 Then
        
        PercentChange = 0
        
        Else
        PercentChange = ws.Cells(i_summary, 10).Value / ws.Cells(i_summary, 13).Value
        ws.Cells(i_summary, 11).Value = PercentChange
        End If
     
    'Reset values for the next loop iteration
        TotalVolume = 0
        PercentChange = 0
        
     'Shift row counter down by 1
        
        i_summary = i_summary + 1
    
        Else
    
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        End If

    Next i

'Populate values for summary table

    For j = 2 To lastrow

'If Statement to detemine the greatest stock increase
        If ws.Cells(j, 11).Value > TopStockChange Then
        TopStockChange = ws.Cells(j, 11).Value
        TopStock = ws.Cells(j, 9).Value
        End If
    
 'If Statement to determine the worst performing stock
    
        If ws.Cells(j, 11).Value < WorstStockChange Then
        WorstStockChange = ws.Cells(j, 11).Value
        WorstStock = ws.Cells(j, 9).Value
        End If
    
'If Statement to determine highest total stock volume
    
        If ws.Cells(j, 12).Value > MostStockVolumeAmt Then
        MostStockVolumeAmt = ws.Cells(j, 12).Value
        MostStockVolume = ws.Cells(j, 9).Value
        End If
    
    Next j
        
'Populate Summary Table

        ws.Range("Q2").Value = TopStock
        ws.Range("R2").Value = TopStockChange
        ws.Range("Q3").Value = WorstStock
        ws.Range("R3").Value = WorstStockChange
        ws.Range("Q4").Value = MostStockVolume
        ws.Range("R4").Value = MostStockVolumeAmt


'Convert Percentage Change Cells to Percent Format

        ws.Range("R2").NumberFormat = "0.00%"
        ws.Range("R3").NumberFormat = "0.00%"
        
Next ws

MsgBox ("Stock Macros Completed for all Worksheets! Thank you Rutgers Data Science Bootcamp!")

End Sub




