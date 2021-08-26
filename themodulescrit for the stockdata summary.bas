Attribute VB_Name = "Module1"
Sub stockdatasummary()

'Declare worksheet variable and loop through the worksheets as active
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate

'initialize ticker variable, yearlychange, percentage change, and volume
Dim ticker As String
Dim yearlychange As Double
Dim percentage As Double
Dim volume As Double
volume = 0

'summary table with headers
Dim summarytable As Integer
summarytable = 2
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
'find the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'declare counter for yearly change
    Dim counter As Integer
    counter = 0
    'open and close variables
    Dim opendate As Double
    Dim closedate As Double

    
'go through the tickers and have it count, if no count then include it into the opendate
For i = 2 To lastrow

    'If counter = 0 Then...
    If counter = 0 Then
    
        opendate = Cells(i, 3).Value
    
    End If
   
   'add to counter for yearly change
        counter = counter + 1
          
    'check throug the ticker name
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set ticker value
    ticker = Cells(i, 1).Value
      
    'close date
    closedate = Cells(i, 6).Value
    
    'calculate yearly change and percent change
    yearlychange = closedate - opendate

    If opendate <> 0 Then
        percentage = yearlychange / opendate
    
    End If
    
    'include it to the initial volume
    volume = volume + Cells(i, 7).Value
    
    'Output the ticker, yearly change with color formats
    Range("I" & summarytable).Value = ticker

    Range("J" & summarytable).Value = yearlychange
    
        If Range("J" & summarytable).Value < 0 Then
            Range("J" & summarytable).Interior.ColorIndex = 23
            
        Else
            Range("J" & summarytable).Interior.ColorIndex = 26
        
        End If
    
    'So outpout to the summary percentage, volume, and add a new row and reset the counters and volume if its the last ticker with that name
    Range("K" & summarytable).Value = percentage
    
    Range("L" & summarytable).Value = volume
    
    summarytable = summarytable + 1
        
    counter = 0
      
    volume = 0
    
'the next ticker is the same then just update the volume as before

    Else
    
volume = volume + Cells(i, 7).Value
    
    End If
'on to the next worksheet and will go all the steps as before
Next i
    
  Next ws
    

End Sub
