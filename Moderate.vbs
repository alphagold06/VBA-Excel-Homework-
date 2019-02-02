Option Explicit

Sub stockmoderate()

Dim Ticker As String
Dim volumetotal As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Summary_Table_Row As Integer
Dim openingprice As Double
Dim closingprice As Double
Dim LastRow As Double
Dim i As Double
Dim totalrows As Long
    
    Summary_Table_Row = 2
    volumetotal = 0
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    totalrows = Application.CountA(Columns(1))
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
      
    Next i
End Sub