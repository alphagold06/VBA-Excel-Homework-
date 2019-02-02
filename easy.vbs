Option Explicit

Sub stockeasy()
Dim Ticker As String
Dim volume As Long
Dim volumetotal As Double
Dim totalrows As Long
Dim totalcols As Long
Dim r As Long
Dim Summary_Table_Row As Integer


volumetotal = 0
Summary_Table_Row = 2
totalrows = Application.CountA(Columns(1))


For r = 2 To totalrows

    'Check if we are still within the ticker, if it is not...
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        
    Ticker = Cells(r, 1).Value
    volumetotal = volumetotal + Cells(r, 7).Value
    
    ' Print the Ticker in the Summary Table
        Range("I1") = "Ticker"
        Range("I" & Summary_Table_Row).Value = Ticker
     

     ' Print the Total Stock Value the Summary Table
        Range("J1") = "Total Stock Value"
        Range("J" & Summary_Table_Row).Value = volumetotal

     ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
     ' Reset the Total Stock Value
        volumetotal = 0

    ' If the cell immediately following a row is the same...
     
     Else

    ' Add to the Total Stock Value
        volumetotal = volumetotal + Cells(r, 7).Value

    End If

  Next r

End Sub