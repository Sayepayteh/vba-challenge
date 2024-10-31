Attribute VB_Name = "Module1"
Sub Quarter1()
    Dim i As Long
    
    Dim ws As Worksheet
    Dim sheetNames As Variant
    
    Dim primTicker As String
    Dim ticker As String
    
    Dim countTicker As Long
    
    Dim dateArray() As String
    Dim yearPart As Integer
    Dim monthPart As Integer
    Dim dayPart As Integer
    
    Dim tickerTitle As String
    Dim quarterlyChangeTitle As String
    Dim percentChangeTitle As String
    Dim totalStockVolTitle As String
    
    Dim greatestIncrTitle As String
    Dim greatestDecrTitle As String
    Dim valueTitle As String
    Dim greatestTotalVolTitle As String
    
    Dim greatestIncr As Double
    Dim greatestDecr As Double
    Dim greatestTotalVol As Double
    
    Dim greatestIncrTicker As String
    Dim greatesDecrTicker As String
    Dim greatestTotalVolTicker As String
    
    
    Dim openVal As Double
    Dim closeVal As Double
    Dim quarterlyChange As Double
    
    Dim primVol As Double
    Dim vol As Double
    
    Dim lastRow As Long
    
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    tickerTitle = "Ticker"
    quarterlyChangeTitle = "Quarterly Change"
    percentChangeTitle = "Percent Change"
    totalStockVolTitle = "Total Stock Volume"
    
    greatestIncrTitle = "Greatest % Increase"
    greatestDecrTitle = "Greatest % Decrease"
    valueTitle = "Value"
    greatestTotalVolTitle = "Greatest Total Volume"
    
    greatestIncr = 0#
    greatestDecr = 0#
    greatestTotalVol = 0#
    
    
    
    For Each sheetName In sheetNames
        
        primVol = 0
        vol = 0
        primTicker = ""
        countTicker = 1
        
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Find the last non-empty row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        
        For i = 2 To lastRow
            ticker = ws.Range("A" & i).Value
            
            If primTicker = "" Then
                ws.Range("I1").Value = tickerTitle
                ws.Range("J1").Value = quarterlyChangeTitle
                ws.Range("K1").Value = percentChangeTitle
                ws.Range("L1").Value = totalStockVolTitle
                
                ws.Range("P2").Value = tickerTitle
                ws.Range("Q2").Value = valueTitle
                
                ws.Range("O3").Value = greatestIncrTitle
                ws.Range("O4").Value = greatestDecrTitle
                ws.Range("O5").Value = greatestTotalVolTitle
            End If
            
            If ticker <> primTicker Then
                primTicker = ticker
                countTicker = countTicker + 1
                
                ws.Range("I" & countTicker).Value = primTicker
                primVol = 0
            End If
            
            vol = CDbl(ws.Range("G" & i).Value)
            primVol = primVol + vol
            
            If vol > greatestTotalVol Then
                greatestTotalVol = vol
                greatestTotalVolTicker = ticker
            End If
            
            dateArray = Split(ws.Range("B" & i).Value, "/")
            yearPart = CInt(dateArray(2))
            monthPart = CInt(dateArray(0))
            dayPart = CInt(dateArray(1))
            
             If (monthPart = 1 And dayPart = 2) Or (monthPart = 4 And dayPart = 1) Or (monthPart = 7 And dayPart = 1) Or (monthPart = 10 And dayPart = 1) Then
                openVal = CDbl(ws.Range("C" & i).Value)
            End If
            
            If (monthPart = 3 And dayPart = 31) Or (monthPart = 6 And dayPart = 30) Or (monthPart = 9 And dayPart = 30) Or (monthPart = 12 And dayPart = 31) Then
                closeVal = CDbl(ws.Range("F" & i).Value)
                
                ' Calculate quarterlyChange only if openVal is not zero
                If openVal <> 0 Then
                    quarterlyChange = closeVal - openVal
                    percentChange = (quarterlyChange / openVal) * 100
                Else
                    quarterlyChange = 0
                    percentChange = 0
                End If
                
                ' Output values to worksheet
                
                ws.Range("J" & countTicker).Value = quarterlyChange
                
                If quarterlyChange < 0 Then
                    ws.Range("J" & countTicker).Interior.Color = RGB(255, 0, 0)
                End If
                
                If quarterlyChange > 0 Then
                    ws.Range("J" & countTicker).Interior.Color = RGB(0, 255, 0)
                End If
                
                
                
                ws.Range("K" & countTicker).Value = Format(percentChange, "0.00") & "%"
                
                If percentChange > greatestIncr Then
                    greatestIncr = percentChange
                    greatestIncrTicker = ticker
                End If
                
                If percentChange < greatestDecr Then
                    greatestDecr = percentChange
                    greatestDecrTicker = ticker
                End If
                
                ' Accumulate vol to primVol
                ws.Range("L" & countTicker).Value = primVol
            End If
        
       
        
        
        Next i
        
        
        
        
    Next sheetName
    
    
    For Each sheetName In sheetNames
        
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        
        ws.Range("Q3").Value = Format(greatestIncr, "0.00") & "%"
        ws.Range("P3").Value = greatestIncrTicker
        
        ws.Range("Q4").Value = Format(greatestDecr, "0.00") & "%"
        ws.Range("P4").Value = greatestDecrTicker
        
        
        ws.Range("Q5").Value = greatestTotalVol
        ws.Range("P5").Value = greatestTotalVolTicker
        
    Next sheetName
    
    
    
End Sub




