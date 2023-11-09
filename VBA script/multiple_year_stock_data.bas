Attribute VB_Name = "Module1"
Sub Stats()

'loop for every worksheet
For Each ws In Worksheets

    'headings of new columns
    ws.Cells(1, 9).Value = ws.Cells(1, 1).Value                   'i
    ws.Cells(1, 10).Value = "Yearly Change"                 'j
    ws.Cells(1, 11).Value = "Percent Change"               'k
    ws.Cells(1, 12).Value = "Total Stock Volume"         'l
    
    'determine last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'give variable for change in row number
    Dim i As Double
    
    'variable for ticker
    Dim ticker As String
    
    'variables of total stock vol
    Dim totalstockvol As Double
    totalstockvol = 0
    
    'tracking of ticker in summary table
    Dim totalvolsumm As Long
    totalvolsumm = 2
    
    'variable for open amount for ticker
    Dim opening As Double
    
    'variable for close amount for ticker
    Dim closeamount As Double
    
    'variable for yearly change
    Dim YC As Double
    
    'variable for percentage change
    Dim percent As Double
    
    opening = ws.Cells(2, 3).Value
    
    'loop through tickers
    For i = 2 To lastrow
    
    
        'if ticker changes
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'set ticker name
        ticker = ws.Cells(i, 1).Value
        
        'add to ticker total
        totalstockvol = totalstockvol + ws.Cells(i, 7).Value
        
        'print ticker name to summary table
        ws.Range("I" & totalvolsumm).Value = ticker
        
        'print total stock volume
        ws.Range("L" & totalvolsumm).Value = totalstockvol
        
        'reset total stock volume
        totalstockvol = 0
        
        'set closing
        closeamount = ws.Cells(i, 6).Value
        
        'calculate yearlychange
        YC = closeamount - opening
        
        'print yearly change
        ws.Range("J" & totalvolsumm).Value = YC
        ws.Range("J" & totalvolsumm).NumberFormat = "0.00"
        
        'change cell color
            If YC >= 0 Then
                ws.Range("J" & totalvolsumm).Interior.ColorIndex = 4
                
            Else
                ws.Range("J" & totalvolsumm).Interior.ColorIndex = 3
            
            End If
        'calculate percentage change
        percent = (YC / opening)
        
        'change opening value
         opening = ws.Cells(i + 1, 3).Value
        
        'print percentage change and changecell type to percentage
        ws.Range("K" & totalvolsumm).Value = percent
        ws.Range("K" & totalvolsumm).Style = "Percent"
        ws.Range("K" & totalvolsumm).NumberFormat = "0.00%"
        
            
         'add one row to summary table
        totalvolsumm = totalvolsumm + 1
        
        'if ticker is the same in cell after
        Else
            
            'add to total volume
            totalstockvol = totalstockvol + ws.Cells(i, 7).Value
            
            
        End If
        
  Next i



'----------------------------------------------------------------------------------
'Determining the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'headings for new summary table

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'define variable for change in row number
Dim j As Double

'define variable for unknowns in summary table
Dim Greater As Double             'greatest percentage increase
Dim lesser As Double                'greatest percentage decrease
Dim vol As Double                       'greatest total volume
Dim tickergreat As String
Dim tickerless As String
Dim tickervol As String



'loop for greateast increase percentage
    For j = 2 To lastrow
    
        If ws.Cells(j, 11).Value > Greater Then
            Greater = ws.Cells(j, 11).Value
            tickergreat = ws.Cells(j, 9).Value
            End If
         
      Next j
      
      'print values in summary table
    ws.Range("q2").Value = Greater
    ws.Range("q2").Style = "Percent"
    ws.Range("p2").Value = tickergreat
    ws.Range("q2").NumberFormat = "0.00%"
    
    'loop for greatest decrease percentage
    For j = 2 To lastrow
    
        If ws.Cells(j, 11).Value < lesser Then
            lesser = ws.Cells(j, 11).Value
            tickerless = ws.Cells(j, 9).Value
            End If
         
    Next j
    
    'print values in summary table
    ws.Range("q3").Value = lesser
    ws.Range("q3").Style = "Percent"
    ws.Range("p3").Value = tickerless
    ws.Range("q3").NumberFormat = "0.00%"
    
    'loop for greatest total volume
    For j = 2 To lastrow
    
        If ws.Cells(j, 12).Value > vol Then
            vol = ws.Cells(j, 12).Value
            tickervol = ws.Cells(j, 9).Value
            End If
         
      Next j
      
      'print values in summary table
    ws.Range("q4").Value = vol
    ws.Range("p4").Value = tickervol
    
 Next ws
    
End Sub



