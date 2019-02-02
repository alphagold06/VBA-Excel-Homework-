Option Explicit

Sub challenge_WS()

Dim Ticker As String
Dim volumetotal As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Summary_Table_Row As Integer
Dim openingprice As Double
Dim closingprice As Double
Dim i As Double
Dim rng1 As Range
Dim rng As Range
Dim FndRng As Range
Dim tickermin, tickermax, tickertotal As String
Dim r As Double
Dim GreatestPerIncrease As Double
Dim GreatestPerDecrease As Double
Dim GreatestTotalVolume As Double
Dim LastRow As Double
Dim totalrows As Long
Dim w as workseet


    Summary_Table_Row = 2
    volumetotal = 0
    totalrows = w.Application.CountA(Columns(1))
    LastRow = w.Cells(Rows.Count, 1).End(xlUp).Row
    openingprice = w.Cells(2, 3)
    
    For Each w in ThisWorkbook.Worksheets
     
     For i = 2 To totalrows
     
     'Check if we are still within the ticker, if it is not...
        If w.Cells(i + 1, 1).Value <> w.Cells(i, 1).Value Then
             
            Ticker = w.Cells(i, 1)
            
            volumetotal = volumetotal + w.Cells(i, 7)
            
            closingprice = w.Cells(i, 6)
            
           'Calc Yearly Change
            YearlyChange = closingprice - openingprice
            
           'Cal percent change if percent change is not 0 
            If openingprice = 0 Then
                PercentChange = 0
            Else
            'Calc Percentage Change
                PercentChange = YearlyChange / openingprice * 100
           End If
           
            w.Range("I" & Summary_Table_Row).Value = Ticker
            w.Range("I1") = "Ticker"
            w.Range("J" & Summary_Table_Row).Value = YearlyChange
            w.Range("J1") = "Yearly Change"
            w.Range("K" & Summary_Table_Row).Value = (PercentChange & "%")
            w.Range("K1") = "Percent Change"
            w.Range("L" & Summary_Table_Row).Value = volumetotal
            w.Range("L1") = "Total Stock Volume"
                             
            volumetotal = 0
       
            openingprice = w.Cells(i + 1, 3)
            
            
             ' Conditional formatting that will highlight positive change in green                         
            If w.Range("J" & Summary_Table_Row).Value >= 0 Then
                w.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
            Else

            'Conditional formatting that will highlight positive change in red
              w.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
                Summary_Table_Row = Summary_Table_Row + 1
        Else
            
                volumetotal = volumetotal + w.Cells(i, 7)
            
        End If
        
'stock with the greatest increase, decrease and total volume

 Set rng = w.Columns(11)
    GreatestPerIncrease = w.Application.max(rng)
    w.Range("Q2") = GreatestPerIncrease * 100
    
    GreatestPerDecrease = w.Application.Min(rng)
    w.Range("Q3") = GreatestPerDecrease * 100
    
    Set rng1 = w.Columns(12)
    GreatestTotalVolume = w.Application.max(rng1)
    w.Range("Q4") = GreatestTotalVolume * 100
           
    totalrows = w.Cells(Rows.Count, 11).End(xlUp).Row
    
For r = 2 To totalrows
        If w.Cells(r, 11) = GreatestPerDecrease Then
            tickermin = w.Cells(r, 9)
        End If
        
        If w.Cells(r, 11) = GreatestPerIncrease Then
            tickermax = w.Cells(r, 9)
            
        End If
            
        If w.Cells(r, 12) = GreatestTotalVolume Then
            tickertotal = w.Cells(r, 9)
            
        End If
            
  Next r
  'column headers
    w.Range("O2") = "Greatest % Increase"
    w.Range("O3") = "Greatest % Decrease"
    w.Range("O4") = "Greatest Total Volume"

  'column values
    w.Range("P1") = "Ticker"
    w.Range("P2") = tickermax
    w.Range("P3") = tickermin
    w.Range("P4") = tickertotal
    
    w.Range("Q1") = "Value"
     
     Next w
    
  Next i
 
    
 
 End Sub