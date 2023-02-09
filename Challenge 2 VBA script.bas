Attribute VB_Name = "Module1"

Sub stockmarket()

Dim A As Integer
Dim ws_count As Integer
Dim starting_worksheet As Worksheet


'setting the worksheet count
ws_count = ActiveWorkbook.Worksheets.Count


'loop through worksheet
For A = 1 To ws_count

ThisWorkbook.Worksheets(A).Activate

' Declaring variables for stockmarket.
    
    Dim i As Long
    Dim lastrow As Long
    Dim counter As Long
    Dim add As Double
    Dim yearlyChange As Double
    Dim pricelow As Double
    Dim pricehigh As Double
    Dim volMax As Double
    Dim pricechange As Boolean
    Dim pricelowTicker As String
    Dim pricehighTicker As String
    Dim volMaxTicker As String
    
    
  ' Filling in the headers.
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
' Initializing variables.

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    counter = 2
    add = 0
    pricechange = True
    pricelow = 1E+99
    pricehigh = -1E+99
    volMax = -1E+99
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Save unique ticker symbol in column I.
            Cells(counter, 9).Value = Cells(i, 1).Value
            
            ' Calculating Yearly Change.
            closePrice = Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            Cells(counter, 10).Value = yearlyChange
            If yearlyChange < 0 Then
                Cells(counter, 10).Interior.Color = vbRed
                
            ElseIf yearlyChange > 0 Then
                Cells(counter, 10).Interior.Color = vbGreen
                
            End If
            
           ' Calculating percent change.
            If yearlyChange = 0 Or openPrice = 0 Then
                Cells(counter, 11).Value = 0
            Else
                Cells(counter, 11).Value = Format(yearlyChange / openPrice, "0.00%")
            End If
        
    
      ' Saving Total Volume into column L.
            add = add + Cells(i, 7).Value
            add = Cells(counter, 12).Value
            
            ' Find the values for greatest decrease/increase and greatest total volume.
            If Cells(counter, 11).Value > pricehigh Then
                If Cells(counter, 11).Value = ".%" Then
                Else
                    pricehigh = Cells(counter, 11).Value
                    pricehighTicker = Cells(counter, 9).Value
                End If
                
            ElseIf Cells(counter, 11).Value < pricelow Then
                pricelow = Cells(counter, 11).Value
                pricelowTicker = Cells(counter, 9).Value
            ElseIf Cells(counter, 12).Value > volMax Then
                volMax = Cells(counter, 12).Value
                volMaxTicker = Cells(counter, 9).Value
            End If
            
            
       ' ticker symbols.
    Cells(2, 16).Value = pricehighTicker
    Cells(3, 16).Value = pricelowTicker
    Cells(4, 16).Value = volMaxTicker

    ' Save the values for greatest decrease/increase and greatest total volume.
    Cells(2, 17).Value = Format(pricehigh, "0.00%")
    Cells(3, 17).Value = Format(pricelow, "0.00%")
    Cells(4, 17).Value = volMax
    
   ' Resetting variables and moving to next ticker symbol.
            counter = counter + 1
            add = 0
            pricechange = True
        Else
            ' Using change to save the open price value at the start of the year.
            If pricechange Then
                openPrice = Cells(i, 3).Value
                pricechange = False
            End If
            ' If adjacent ticker symbols are the same, then save volume value.
            add = add + Cells(i, 7).Value
        End If
    Next i
    
    'move to next sheet
    Next A
    
End Sub
