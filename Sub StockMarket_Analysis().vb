Sub StockMarket_Analysis()


'Create variables :

Dim Ticker As String

Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Total_Stock_Volume As Double

Dim year_open As Double

Dim year_close As Double

Dim start_data As Integer
'----------------------------------------

Dim ws As Worksheet

 ' insert the loop for all worksheets

For Each ws In Worksheets

    'Insert a columns
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Insert names for greatest increase,greatest decrease, and  greatest volume into colums

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    'Create for greatest increase, greatest increase, and  greatest volume Ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = decrease
    ws.Range("P4").Value = Greatest
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


    'Insert  intiger for the loop to start
    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0

    'Go to the end of coumn A

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


        For i = 2 To EndRow

            'If Tickersymbol change or not equal to the previous one , then

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value

            'Insert the variable for the next Ticker Alphabet

            previous_i = previous_i + 1

            ' inset the  value  of  the first day open form and last day close of the year in colum 3 and 6

            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value

            ' sum the total stock volume using a colum 7
            
            For j = previous_i To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j


            If year_open = 0 Then

                Percent_Change = year_close

            Else
                Yearly_Change = year_close - year_open

                Percent_Change = Yearly_Change / year_open

            End If
            
         '--------------------------------------------------

            'insert the values in the worksheet summary table

            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change

            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            
            ws.Cells(start_data, 12).Value = Total_Stock_Volume

            'IIn the data  after first row completed go to the next row

            start_data = start_data + 1

            ' insert the variable to zero again and move i to a previos_i
            Yearly_Change = 0
            Total_Stock_Volume = 0
            Percent_Change = 0

            previous_i = i

        End If

'---------------------------------------------------------

    Next i

    'insert the last row of column k

    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'insert variable to initiate the second summery table value

    Increase = 0
    decrease = 0
    Greatest = 0

        'find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow

            'insert previous increment to check
            last_k = k - 1

            'inert current and prevous_k rows for percentage
            
            current_k = ws.Cells(k, 11).Value
            
            prevous_k = ws.Cells(last_k, 11).Value

            'current greatest total volume and prevous_Volume rows
            
            volume = ws.Cells(k, 12).Value
            
            prevous_vol = ws.Cells(last_k, 12).Value

   '--------------------------------------------------

            'Find the increase
            If Increase > current_k And Increase > prevous_k Then

                Increase = Increase

                'insert name for increase percentage

            ElseIf current_k > Increase And current_k > prevous_k Then

                Increase = current_k

                'name for increase percentage
                
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then

                Increase = prevous_k
                
                increase_name = ws.Cells(last_k, 9).Value

            End If

       '--------------------------------------------------
             ' the decrease

            If decrease < current_k And decrease < prevous_k Then

                decrease = decrease

                'insert name for increase percentage

            ElseIf current_k < Increase And current_k < prevous_k Then

                decrease = current_k

                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then

                decrease = prevous_k

                decrease_name = ws.Cells(last_k, 9).Value

            End If
       '--------------------------------------------------
       
           ' The greatest volume

            If Greatest > volume And Greatest > prevous_vol Then

                Greatest = Greatest

                ' name for greatest volume
            
            ElseIf volume > Greatest And volume > prevous_vol Then

                Greatest = volume

                ' Insert name for greatest volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
            
            Greatest = prevous_vol
                
                ' name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k
  '--------------------------------------------------
  
' Conditional formatting columns colors

'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

Next ws

'--------------------------------------------------
End Sub



Sub Alphabetical_testing()

' Alphabetical_testing Macro



End Sub

