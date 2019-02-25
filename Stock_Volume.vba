Sub stock_volume()

    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate

        'Variables to get volume sum
        
        Dim ticker_name As String
        Dim total_volume As Double
        Dim volume_summary_row As Long
        
        'variables for yearly change info
        
        Dim yearly_change As Double
        Dim percentage_change As String
        Dim first_date_row As Long
        Dim first_date_open As Double
        Dim last_date_close As Double
        
        'Beginning Values
        
        total_volume = 0
        volume_summary_row = 2
        first_date_row = 2

        'last row
        
        Dim last_row As Long
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sort Column First
  
        Range("A2:G" & last_row).Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlYes
        
        'Sort Column
  
        Range("A2:G" & last_row).Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes
        
        Dim last_sum_row As Long
        last_sum_row = Cells(Rows.Count, 10).End(xlUp).Row
        
        'Summary Table Headers
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
            
        'Loop through all volumes to get summary of total volume by ticker

        For i = 2 To last_row

            'check if the ticker is the same, if it's not...
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                
                'Insert summary values
                
                ticker_name = Cells(i, 1).Value
                total_volume = total_volume + Cells(i, 7).Value
                first_date_open = Cells(first_date_row, 3).Value
                last_date_close = Cells(i, 6).Value
                yearly_change = last_date_close - first_date_open
                'check if divisor is 0 and if not print percentage change
                    If first_date_open = 0 Then
                        percentage_change = 0
                    Else: percentage_change = FormatPercent(yearly_change / first_date_open, 2)
                    End If
                
                Range("I" & volume_summary_row).Value = ticker_name
                Range("J" & volume_summary_row).Value = yearly_change
                Range("K" & volume_summary_row).Value = percentage_change
                Range("L" & volume_summary_row).Value = total_volume
                
                'Adjust colors
                
                    If Range("K" & volume_summary_row).Value > 0 Then
                       Range("K" & volume_summary_row).Interior.ColorIndex = 4
                    Else: Range("K" & volume_summary_row).Interior.ColorIndex = 3
                    End If
                
                'add one to the summary row'
                volume_summary_row = volume_summary_row + 1
    
                'reset
                total_volume = 0
                first_date_row = i + 1
            
            'If same, add to total volume'
            Else: total_volume = total_volume + Cells(i, 7).Value
                
            End If
              
        Next i
        
        'Max and mins from summary table
        
        max_increase = WorksheetFunction.Max(ws.Range("K2:K" & last_sum_row))
        inc_ticker = WorksheetFunction.Match(max_increase, ws.Range("K2:K" & last_sum_row), 0)
        max_decrease = WorksheetFunction.Min(ws.Range("K2:K" & last_sum_row))
        dec_ticker = WorksheetFunction.Match(max_decrease, ws.Range("K2:K" & last_sum_row), 0)
        max_volume = WorksheetFunction.Max(ws.Range("L2:L" & last_sum_row))
        max_vol_ticker = WorksheetFunction.Match(max_volume, ws.Range("L2:L" & last_sum_row), 0)
        
        Range("Q2").Value = FormatPercent(max_increase, 2)
        Range("P2").Value = Cells(inc_ticker + 1, 9)
        Range("Q3").Value = FormatPercent(max_decrease, 2)
        Range("P3").Value = Cells(dec_ticker + 1, 9)
        Range("Q4").Value = max_volume
        Range("P4").Value = Cells(max_vol_ticker + 1, 9)
        
        
        
        
    Next ws
    
End Sub