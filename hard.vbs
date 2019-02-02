Option Explicit

Sub stockhard()

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


    Summary_Table_Row = 2
    volumetotal = 0
    totalrows = Application.CountA(Columns(1))
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    openingprice = Cells(2, 3)
    
    
     
     For i = 2 To totalrows
     
     'Check if we are still within the ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             
            Ticker = Cells(i, 1)
            
            volumetotal = volumetotal + Cells(i, 7)
            
            closingprice = Cells(i, 6)
            
           'Calc Yearly Change
            YearlyChange = closingprice - openingprice
            
           'Cal percent change if percent change is not 0 
            If openingprice = 0 Then
                PercentChange = 0
            Else
            'Calc Percentage Change
                PercentChange = YearlyChange / openingprice * 100
           End If
           
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("I1") = "Ticker"
            Range("J" & Summary_Table_Row).Value = YearlyChange
            Range("J1") = "Yearly Change"
            Range("K" & Summary_Table_Row).Value = (PercentChange & "%")
            Range("K1") = "Percent Change"
            Range("L" & Summary_Table_Row).Value = volumetotal
            Range("L1") = "Total Stock Volume"
                             
            volumetotal = 0
       
            openingprice = Cells(i + 1, 3)
            
            
             ' Conditional formatting that will highlight positive change in green                         
            If Range("J" & Summary_Table_Row).Value >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
            Else

            'Conditional formatting that will highlight positive change in red
              Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
                Summary_Table_Row = Summary_Table_Row + 1
        Else
            
                volumetotal = volumetotal + Cells(i, 7)
            
        End If
        
'stock with the greatest increase, decrease and total volume

 Set rng = Columns(11)
    GreatestPerIncrease = Application.max(rng)
    Range("Q2") = GreatestPerIncrease * 100
    
    GreatestPerDecrease = Application.Min(rng)
    Range("Q3") = GreatestPerDecrease * 100
    
    Set rng1 = Columns(12)
    GreatestTotalVolume = Application.max(rng1)
    Range("Q4") = GreatestTotalVolume * 100
           
    totalrows = Cells(Rows.Count, 11).End(xlUp).Row
    
For r = 2 To totalrows
        If Cells(r, 11) = GreatestPerDecrease Then
            tickermin = Cells(r, 9)
        End If
        
        If Cells(r, 11) = GreatestPerIncrease Then
            tickermax = Cells(r, 9)
            
        End If
            
        If Cells(r, 12) = GreatestTotalVolume Then
            tickertotal = Cells(r, 9)
            
        End If
            
  Next r
  'column headers
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"

  'column values
    Range("P1") = "Ticker"
    Range("P2") = tickermax
    Range("P3") = tickermin
    Range("P4") = tickertotal
    
    Range("Q1") = "Value"
     
    
  Next i
 
    
 
 End Sub