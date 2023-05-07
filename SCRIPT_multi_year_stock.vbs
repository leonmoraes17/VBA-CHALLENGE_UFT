Sub unique_ticker()

Dim ws As Worksheet

For Each ws In Worksheets

Dim ticker_name As String
Dim ticker_summary As Integer

Dim yearly_summary As Integer
Dim opening_value As Double
Dim closing_value As Double

Dim percentage_change As Double
Dim percentage_summary As Integer

Dim total_volume As LongLong
Dim volume_summary As Integer

Dim lastrow As Long

Dim yearly_change As Double

volume_summary = 2
percentage_summary = 2
yearly_summary = 2
ticker_summary = 2



    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    ws.Range("I1").Value = "Unique_Ticker"
    ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = " Percentage Change'"
    ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker_name = ws.Cells(i, 1).Value
            ws.Range("I" & ticker_summary) = ticker_name
            ticker_summary = ticker_summary + 1
            End If
        Next i
    'yearly change module
    
    opening_value = Cells(2, 3).Value
        For j = 2 To lastrow
        
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then 'j3<>j2
            'yearly price change
                yearly_change = ws.Cells(j, 6).Value - opening_value
                ws.Range("J" & yearly_summary).Value = yearly_change
                yearly_summary = yearly_summary + 1
            
            'yearly percentage change
            If yearly_change = 0 Then
                percentage_change = 0
            Else: percentage_change = (yearly_change / opening_value) * 100
            End If
            
            ws.Range("K" & percentage_summary).Value = percentage_change
            percentage_summary = percentage_summary + 1  'can use the variable yearly_summary as well as the integer value is identical,
            'however for clarity purpose as new to programming , i am putting a new variable
             
            'TOAL STOCK VOLUME
                       
            
            opening_value = ws.Cells(j + 1, 3).Value
            'basically hardcoding the first opening balance of ticker and then taking the value of next ticker thereafter
            'so if j3=j2 then no action is taken and value of j is incremented by 1 and if j3 <> j2 then the if loop is generated
            
            End If
            'yearly_change = 0
            
            Next j
          'changing the cell color
          For w = 2 To lastrow
            If ws.Cells(w, 10).Value <= 0 Then
                ws.Cells(w, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(w, 10).Interior.ColorIndex = 4
                End If
            Next w
            
          'getting the total volume of each ticker
          Dim l As Integer
                                 
          
          For l = 2 To lastrow
            If ws.Cells(l + 1, 1) <> ws.Cells(l, 1) Then
                total_volume = total_volume + ws.Cells(l, 7).Value
                ws.Range("L" & volume_summary) = total_volume
                volume_summary = volume_summary + 1
                total_volume = 0
            Else
                total_volume = total_volume + ws.Cells(l, 7).Value
            End If
        Next l
        
        
'adding functionality

ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest total volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"


'Max value

Dim maxvalue As Double
Dim max_ticker As String

Dim x As Long
lastrow1 = ws.Range("K" & Rows.count).End(xlUp).Row

    
    maxvalue = 0
    
For x = 2 To lastrow1
    If ws.Range("K" & x).Value > maxvalue Then
        maxvalue = ws.Range("K" & x).Value
        ws.Range("R2") = maxvalue
       
       max_ticker = ws.Range("I" & x)
       ws.Range("q2").Value = max_ticker
    End If
    Next x
    
    
 'MIN VALUE

Dim minvalue As Double
Dim min_ticker As String

Dim y As Long
lastrow1 = ws.Range("K" & Rows.count).End(xlUp).Row

    
    minvalue = 0
    
For y = 2 To lastrow1
    If ws.Range("K" & y).Value < minvalue Then
        minvalue = ws.Range("K" & y).Value
        ws.Range("R3") = minvalue
       
       min_ticker = ws.Range("I" & y)
       ws.Range("Q3").Value = min_ticker
    End If
    Next y
    
'total VOLUME

Dim totalvalue As LongLong

Dim volume_ticker As String

Dim z As Long
lastrow2 = ws.Range("L" & Rows.count).End(xlUp).Row

    
    totalvolume = 0
    
For z = 2 To lastrow2
    If ws.Range("L" & z).Value > totalvolume Then
        totalvolume = ws.Range("L" & z).Value
        ws.Range("R4").Value = totalvolume
       
            
       volume_ticker = ws.Range("I" & z)
       ws.Range("Q4").Value = volume_ticker
       ws.Range("r4").NumberFormat = "#,##0.00??"
    End If
    Next z
        
Next ws

End Sub


