Attribute VB_Name = "Module2"
Sub StockMarket_Analysis()

'Need to sort Spreadsheet based on Ticker. Otherwise this code will not work.

Dim Ticker As String
Dim Opening As Double
Dim Closing As Double
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double
Dim start_data As Integer
Dim ws As Worksheet
Dim Stock_Volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_Total_volume As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_total_volume_ticker As String



For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Opening"
    ws.Range("K1").Value = "Closing"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "greatest_percent_increase"
    ws.Range("P3").Value = "greatest_percent_decrease"
    ws.Range("P4").Value = "greatest_total_volume"
    
    First_calculation_counter = 2
    
    start_data = 2
    end_data = 1
    Total_Stock_Volume = 0
    
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    '-------------------------------------------------------------------------------------------------------------------------------
    'Beginning of First set of Calculation ( Unique Ticket , opening , closing , yearly change , percentage change and Total Volume)
    'This Calculation will populate Column i to n in same Worksheet
    
    
    
    For i = 2 To RowCount
      
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
          Ticker = ws.Cells(i, 1).Value
          ws.Cells(First_calculation_counter, 9).Value = Ticker
          Opening = ws.Cells(start_data, 3)
          ws.Cells(First_calculation_counter, 10).Value = Opening
          Closing = ws.Cells(i, 6)
          ws.Cells(First_calculation_counter, 11) = Closing
          Yearly_Change = Closing - Opening
          Percent_Change = ((Yearly_Chage / Opening) * 100)
          ws.Cells(First_calculation_counter, 12) = Yearly_Change
          Percent_Change = (Yearly_Change / Opening)
          ws.Cells(First_calculation_counter, 13) = Percent_Change
          ws.Cells(First_calculation_counter, 13).NumberFormat = "0.00%"
          Stock_Volume = ws.Cells(i, 7).Value
          Total_Stock_Volume = Total_Stock_Volume + Stock_Volume
          ws.Cells(First_calculation_counter, 14) = Total_Stock_Volume
          
          end_data = i
          Total_Stock_Volume = 0
          
          First_calculation_counter = First_calculation_counter + 1
          start_data = i + 1
      Else
          Stock_Volume = ws.Cells(i, 7).Value
          Total_Stock_Volume = Total_Stock_Volume + Stock_Volume
          
      End If
      
    
    Next i
    
    'End of First set of Calculation ( Unique Ticket , opening , closing , yearly change , percentage change and Total Volume)

    
'-------------------------------------------------------------------------------------------------------------------------------
' Condistional formatting

'The end row for column I


    Endrow_first_set_of_Calculation = ws.Cells(Rows.Count, "I").End(xlUp).Row


        For j = 2 To Endrow_first_set_of_Calculation

            'if greater than or less than zero
            If ws.Cells(j, 12) > 0 Then

                ws.Cells(j, 9).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 9).Interior.ColorIndex = 3
            End If

        Next j
    

'-------------------------------------------------------------------------------------------------------------------------------

'Calculation of greatest percent increase , grestest percent decrease and greatest volume
'for clarity doing this calculation using another loop otherwise it can be combined with previous conditional formatting loop

 temp_total_volume = 0
 temp_greatest_increase = 0
 temp_greatest_decrease = 0
 
 For k = 2 To Endrow_first_set_of_Calculation
      
            
      If ws.Cells(k, 13) > temp_greatest_increase Then
               
         temp_greatest_increase = ws.Cells(k, 13)
         greatest_percent_increase_ticker = ws.Cells(k, 9)
         
      End If
      
      If ws.Cells(k, 13) < temp_greatest_decrease Then
               
         temp_greatest_decrease = ws.Cells(k, 13)
         greatest_percent_decrease_ticker = ws.Cells(k, 9)
         
      End If
      
      If ws.Cells(k, 14) > temp_total_volume Then
         
         temp_total_volume = ws.Cells(k, 14)
         greatest_total_volume_ticker = ws.Cells(k, 9)
         
      End If
      
 Next k
 
  
 greatest_percent_increase = temp_greatest_increase
 ws.Range("Q2").Value = greatest_percent_increase_ticker
 ws.Range("R2").Value = greatest_percent_increase
 ws.Range("R2").NumberFormat = "0.00%"
 
 greatest_percent_decrease = temp_greatest_decrease
 ws.Range("Q3").Value = greatest_percent_decrease_ticker
 ws.Range("R3").Value = greatest_percent_decrease
 ws.Range("R3").NumberFormat = "0.00%"
 
 
 greatest_Total_volume = temp_total_volume
 ws.Range("Q4").Value = greatest_total_volume_ticker
 ws.Range("R4").Value = greatest_Total_volume
 

'-------------------------------------------------------------------------------------------------------------------------------

'Next Worksheet

Next ws



End Sub
