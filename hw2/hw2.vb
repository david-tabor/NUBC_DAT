Sub hw2()
    Dim i As Double
    Dim i_out As Integer
    Dim sum As Double
    Dim annual_start_price As Double
    Dim annual_end_price As Double
    Dim ticker As String
    Dim nextTicker As String
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Activate
      
      i = 2
      i_out = 2
      sum = 0
      annual_start_price = Cells(2, 3)
      
      Do While IsEmpty(Cells(i, 1)) = False
          ticker = Cells(i, 1)
          nextTicker = Cells(i + 1, 1)
          
          If (ticker = nextTicker) Then
              sum = sum + Cells(i, 7)
          Else
              'Have reached end of this ticker, close price is annual end price
              annual_end_price = Cells(i, 6)
              
              'Write output and increment output row
              Cells(i_out, 9) = ticker
              Cells(i_out, 12) = sum + Cells(i, 7)
              Cells(i_out, 10) = annual_end_price - annual_start_price
              
              If (annual_start_price > 0) Then
                  Cells(i_out, 11) = (annual_end_price - annual_start_price) / annual_start_price
              End If
              
              i_out = i_out + 1
              
              'Initialize variables for next ticker name
              sum = 0
              annual_start_price = Cells(i + 1, 3)
          End If
          
          i = i + 1
      Loop
      
      'Apply formatting to primary output table
      Range("I1") = "Ticker"
      Range("J1") = "Yearly Change"
      With Range("J2:J" & i_out)
          .NumberFormat = "$0.00"
          .FormatConditions.Delete
          .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
          .FormatConditions(1).Interior.Color = RGB(0, 255, 0)
          .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
          .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
      End With
      Range("K1") = "Percent Change"
      Range("K2:K" & i_out).NumberFormat = "0.00%"
      Range("L1") = "Total Stock Volume"
      
      'Hard Challenge - Initialize secondary output table
      Range("O2") = "Greatest % Increase"
      Range("O3") = "Greatest % Decrease"
      Range("O4") = "Greatest Total Volume"
      Range("P1") = "Ticker"
      Range("Q1") = "Value"
      Range("Q2") = 0
      Range("Q3") = 0
      Range("Q4") = 0
      Range("Q2:Q3").NumberFormat = "0.00%"
      
      'Hard Challenge - Scan primary output table for secondary output
      For i = 2 To i_out
          If (Cells(i, 11) > Range("Q2")) Then
              Range("P2") = Cells(i, 9)
              Range("Q2") = Cells(i, 11)
          ElseIf (Cells(i, 11) < Range("Q3")) Then
              Range("P3") = Cells(i, 9)
              Range("Q3") = Cells(i, 11)
          End If
          
          If (Cells(i, 12) > Range("Q4")) Then
              Range("P4") = Cells(i, 9)
              Range("Q4") = Cells(i, 12)
          End If
      Next i

    Next ws
End Sub

