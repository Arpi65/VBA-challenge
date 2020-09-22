Attribute VB_Name = "Module1"
Sub final():
 For Each ws In Worksheets

Dim Lastrow As Long
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 11).Value = "Tickers"
ws.Cells(1, 12).Value = "Yearly Change"
ws.Cells(1, 13).Value = "Percent Change"
ws.Cells(1, 14).Value = "Total Stock Volume in TH"
Dim tot As Double
Dim ctot As Double
Dim j As Long
Dim k As Long
Dim perc As Double

ctot = 0
j = 2
k = 2


    For i = 2 To Lastrow
    
     
    
     If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
     
        tot = ws.Cells(i, 7).Value
        
         ctot = ctot + tot
         
          k = k + 1
     
     ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
     
     'Get the ticker name
     
     ws.Cells(j, 11).Value = ws.Cells(i, 1)
     
     'Do the math to find the difference for the year
     
     Change = ws.Cells(i, 3) - ws.Cells(i - k + 2, 3)
     
           
     ws.Cells(j, 12).Value = Change
     
     
     'calc the percent change
     'put if function for ws.Cells(i - k + 2, 3) when it's equal to 0
         If ws.Cells(i - k + 2, 3) <> 0 Then
            perc = (((ws.Cells(i, 3) - ws.Cells(i - k + 2, 3)) / ws.Cells(i - k + 2, 3)) * 100)
            ws.Cells(j, 13).Value = perc
            Else: ws.Cells(j, 13).Value = "NA"
         End If
        
     'Put the totals in order
                  
     ws.Cells(j, 14).Value = ctot + ws.Cells(i, 7).Value
     
     'Reset for the next tickerh
        
        ctot = 0
        j = j + 1
        k = 2
         
    End If
    
    
Next i

'Finding max min and max total and formating/second part

Dim Lastr As Long

ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"


Lastr = ws.Cells(Rows.Count, 12).End(xlUp).Row

'Find max of L=price change and find the ticker

ws.Cells(2, 18).Value = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & Lastr))

  For p = 1 To Lastr
    
        If ws.Cells(p, 12).Value = ws.Cells(2, 18).Value Then
        
        
            ws.Cells(2, 17).Value = ws.Cells(p, 11).Value
            
      
        End If
    Next p


'Find minimum of L=price change and find the ticker

ws.Cells(3, 18).Value = WorksheetFunction.Min(ws.Range("L2" & ":" & "L" & Lastr))


    For p = 1 To Lastr
    
    
        If ws.Cells(p, 12).Value = ws.Cells(3, 18).Value Then
        
        
            ws.Cells(3, 17).Value = ws.Cells(p, 11).Value
            
      
        End If
        
        If ws.Cells(p, 12).Value >= 0 Then
        ws.Range("L" & p).Interior.ColorIndex = 3
        Else: ws.Range("L" & p).Interior.ColorIndex = 4
        End If
        
    Next p

'Find maximum of N=volume and find the ticker
ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("N2" & ":" & "N" & Lastr))

 For p = 1 To Lastr
    
        If ws.Cells(p, 14).Value = ws.Cells(4, 18).Value Then
        
        
            ws.Cells(4, 17).Value = ws.Cells(p, 11).Value
            
      
        End If
    Next p



'MsgBox (ctot)

MsgBox (Lastrow)

Next ws

End Sub


