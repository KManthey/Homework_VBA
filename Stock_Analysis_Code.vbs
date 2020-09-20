Sub Stock_Analysis()
'VBA Multiple Year Stock Data Analysis
'activate the entire workbook for loops
For Each WS In ActiveWorkbook.Worksheets
WS.Activate


'declare variables and set initial values
Dim Vol As Double
Vol = 0
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim First_Open As Double
Dim Last_Close As Double
Dim Ticker As String
Dim rowCounter As Long

'determine end of the row - xlUp makes sure we do not include a blank row
EndrowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Sets Start value for the row to account for row headers
rowCounter = 2

'Headers for results columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'loop through from 2 until end
For i = 2 To EndrowCount
    If i = 2 Then
        Ticker = Cells(i, 1).Value
        First_Open = Cells(i, 3).Value
    End If
    'set last close value inside loop
    Last_Close = Cells(i, 6).Value

'if cell values in column A do not match then print ticker and sum of vol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'assign ticker to range
            Range("I" & rowCounter).Value = Ticker
            Ticker = Cells(i + 1, 1).Value
            
            'Calculation and destination for Yearly_Change - rounded
            Yearly_Change = Round((Last_Close - First_Open), 2)
            Range("J" & rowCounter).Value = Yearly_Change
                        
            'Conditional for Percent_Change
            If (Last_Close = 0 And First_Open = 0) Then
                Percent_Change = 0
            ElseIf (First_Open = 0 And Last_Close <> 0) Then
                Percent_Change = 1
            Else
                Percent_Change = (Yearly_Change / First_Open)
                'Test this is the only format I can get to stick here nothing for "0.00% or NumberFormat works! - still gets error
                
            End If
              'Reset values
           First_Open = Cells(i + 1, 3).Value
            
            'assign % change to K column and format
            Range("K" & rowCounter).Value = Percent_Change
            Range("K" & rowCounter).NumberFormat = "0.00%"

            'Sum of Vol and return vol value - this is part of the first If statement
            Vol = Vol + Cells(i, 7).Value
            Range("L" & rowCounter).Value = Vol
            
            'Reset vol after returned - Note Vol is set to continue to count in the last Else statement as part of the first If statement
            Vol = 0
           
        If (Yearly_Change < 0) Then
            Range("J" & rowCounter).Interior.ColorIndex = 3
        Else
            Range("J" & rowCounter).Interior.ColorIndex = 4
        End If
        
          'reset rowCounter
           rowCounter = rowCounter + 1
        Else
            Vol = Vol + Cells(i, 7).Value
            
        End If
        
         Next i

'Final winners determined outside of loop
    'Add Row Headers - in column O
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    'Max function - output to P2 from range K - added format for max and min function
    Range("P2") = WorksheetFunction.Max(Range("K2:K" & EndrowCount))
    Range("P2:P3").NumberFormat = "0.00%"
    
    'TickerNumber associated with Max - calculate and store
    TickerNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & EndrowCount)), Range("K2:K" & EndrowCount), 0)
    'Output TickerNumber
    Range("O2") = Cells(TickerNumber + 1, 9)
    
   'Min function - output to P3 from range K
    Range("P3") = WorksheetFunction.Min(Range("K2:K" & EndrowCount))
    
    'TickerNumber associated with Min - calculate and store
    TickerNumber = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & EndrowCount)), Range("K2:K" & EndrowCount), 0)
    'Output TickerNumber
    Range("O3") = Cells(TickerNumber + 1, 9)
    
    'Max for L - Greatest total
    Range("P4") = WorksheetFunction.Max(Range("L2:L" & EndrowCount))
   'TickerNumber associated with Max - calculate and store
    TickerNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & EndrowCount)), Range("L2:L" & EndrowCount), 0)
    'Output TickerNumber
    Range("O4") = Cells(TickerNumber + 1, 9)
    
   Next WS
         
    End Sub
