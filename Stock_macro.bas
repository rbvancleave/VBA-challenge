Attribute VB_Name = "Stock_Macro"
Sub Get_Ticker():
    
'Execute through each sheet

    For Each ws In Worksheets
    
'Set Counter
        
        counter = 2

'Write Headers

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'Find Last Row

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
' Go throught each row

        For i = 2 To lastRow
            
' Ticker is different from previous row, write Ticker name, save year open variable and add total stock volume
            
           If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                tick = ws.Cells(i, 1).Value
                Year_Open = ws.Cells(i, 3)
                
                ws.Cells(counter, 9).Value = tick
                ws.Cells(counter, 12).Value = ws.Cells(counter, 12) + ws.Cells(i, 7)
                
' This is the Last row for the current ticker
                
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                year_close = ws.Cells(i, 6)
                
                ws.Cells(counter, 10).Value = (year_close - Year_Open)
                
    ' Prevent Year_open 0 error

                If Year_Open = 0 Then
                
                    ws.Cells(counter, 11).Value = 0
                
    ' calculate % change, add vol to total stock volume, and increase counter
                
                Else
                
                    ws.Cells(counter, 11).Value = (year_close - Year_Open) / Year_Open
                
                End If
                
                
                ws.Cells(counter, 12) = ws.Cells(counter, 12) + ws.Cells(i, 7)
                
                counter = counter + 1
                
                
 ' Add vol to total stock volume
 
            Else
                
                ws.Cells(counter, 12) = ws.Cells(counter, 12) + ws.Cells(i, 7)
                
            End If
            
        Next i
        
' Pause to ensure no crashing
        
        Application.Wait (Now + TimeValue("0:00:05"))
        
' Find last row of summary section (I:L)
        
        lastSumRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
' Format Percent Change to %

        ws.Range("K2:K" & lastSumRow).NumberFormat = "0.00%"
        
' Format Cell Color to Green if Change > 0 and Red if Change < 0
        
        For i = 2 To lastSumRow
            If ws.Cells(i, 10) <= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
                
        Next i

' Pause to ensure no crashing

        Application.Wait (Now + TimeValue("0:00:05"))
        
'Bonus: Label & Format cells for Greatest Summary section
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
' Set variables
        Max = 0
        Min = 0
        Maxtick = ""
        Mintick = ""
        Vol_max = 0
        Vol_max_tick = ""
        
' Go through each row and compare metric to the current Max/Min etc.
        For i = 2 To lastSumRow
            If ws.Cells(i, 11) > Max Then
                
                Max = ws.Cells(i, 11).Value
                Maxtick = ws.Cells(i, 9).Value
                
                ws.Cells(2, 17).Value = Max
                ws.Cells(2, 16).Value = Maxtick
            
            End If
        
        Next i
        
        For i = 2 To lastSumRow
            If ws.Cells(i, 11) < Min Then
                
                Min = ws.Cells(i, 11).Value
                Mintick = ws.Cells(i, 9).Value
                
                ws.Cells(3, 17).Value = Min
                ws.Cells(3, 16).Value = Mintick
            
            End If
        
        Next i
        
        For i = 2 To lastSumRow
            If ws.Cells(i, 12) > Vol_max Then
                Vol_max = ws.Cells(i, 12).Value
                Vol_max_tick = ws.Cells(i, 9).Value
                
                ws.Cells(4, 17).Value = Vol_max
                ws.Cells(4, 16).Value = Vol_max_tick
            
            End If
        
        Next i
        
    Next ws

End Sub
