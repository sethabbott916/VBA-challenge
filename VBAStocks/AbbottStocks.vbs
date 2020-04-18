Attribute VB_Name = "Module1"
Sub multi()

For Each ws In Worksheets


Dim oldticker As String
Dim newticker As String
Dim lastrow As Long

Dim yearlychange As Double
Dim percentchange As Double

Dim totalvolume As Double
Dim currentvolume As Double

Dim openprice As Double
Dim closeprice As Double

Dim greatvolume As Double
Dim volumetick As String

Dim greatincrease As Double
Dim increasetick As String
Dim greatdecrease As Double
Dim decreasetick As String


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest total volume"
ws.Cells(1, 16).Value = "Ticker"


oldticker = ws.Cells(2, 1).Value
ws.Cells(2, 9).Value = oldticker

openprice = ws.Cells(2, 3).Value


greatvolume = 0
greatincrease = 0


    For x = 2 To lastrow + 1
    
    
    newticker = ws.Cells(x, 1).Value
    currentvolume = ws.Cells(x, 7).Value
    
        If oldticker = ws.Cells(x, 1).Value Then
        newticker = ws.Cells(x, 1).Value
        oldticker = newticker
        
        
        totalvolume = currentvolume + totalvolume
        
        
        Else
        
        closeprice = ws.Cells((x - 1), 6).Value
        yearlychange = closeprice - openprice

            If openprice = 0 Then
            percentchange = 0

            Else
            percentchange = yearlychange / openprice

            End If


        openprice = ws.Cells(x, 3).Value
        
            For y = 2 To 3169
            
                If ws.Cells(y, 9).Value = oldticker Then
                ws.Cells(y + 1, 9).Value = newticker
                ws.Cells(y, 10).Value = yearlychange
                ws.Cells(y, 11).Value = FormatPercent(percentchange)
                ws.Cells(y, 12).Value = totalvolume
                End If
                
                
                
                If ws.Cells(y, 10).Value > 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(y, 10).Value < 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = 3
                End If
                
                
                
                If totalvolume > greatvolume Then
                greatvolume = totalvolume
                volumetick = oldticker
            
                ElseIf percentchange > greatincrease Then
                greatincrease = percentchange
                increasetick = oldticker
            
                ElseIf percentchange < greatdecrease Then
                greatdecrease = percentchange
                decreasetick = oldticker
                End If
                
            Next y
            
          
        oldticker = newticker

        totalvolume = currentvolume
            
        End If
    
    
    Next x
        
    ws.Cells(2, 16).Value = increasetick
    ws.Cells(2, 17).Value = FormatPercent(greatincrease)
    ws.Cells(3, 16).Value = decreasetick
    ws.Cells(3, 17).Value = FormatPercent(greatdecrease)
    ws.Cells(4, 16).Value = volumetick
    ws.Cells(4, 17).Value = greatvolume
    
Next ws

End Sub

