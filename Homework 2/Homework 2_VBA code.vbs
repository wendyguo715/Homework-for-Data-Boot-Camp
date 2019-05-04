Option Explicit

Sub currentsheet()
Dim i As Long
Dim j As Long
Dim lastrow As Long
Dim totalvolume As Double
Dim yearlyopen As Double
Dim k As Long
Dim klastrow As Long
Dim maxticker As String
Dim minticker As String
Dim maxvolume As String
Dim max2 As Double
Dim max As Double
Dim min As Double
Dim ws As Worksheet

For Each ws In Worksheets

    'start of moderate
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    totalvolume = 0
    yearlyopen = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(j, 10).Value = ws.Cells(i, 6).Value - yearlyopen
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            ws.Cells(j, 12).Value = totalvolume + ws.Cells(i, 7).Value
                'Because there are 0 value under "<open>" column while 0 cannot be the denominator, set up percentage to be 0 while yearlyopen is 0. Confirmed with TA.
                If yearlyopen = 0 Then
                ws.Cells(j, 11).Value = 0
                Else
                ws.Cells(j, 11).Value = Format((ws.Cells(i, 6).Value - yearlyopen) / yearlyopen, "Percent")
                End If
            j = j + 1
            totalvolume = 0
            yearlyopen = ws.Cells(i + 1, 3).Value
        Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
    
    Next i
    
    'Start of Hard
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest % Total Volume"
    max = 0
    min = 0
    max2 = 0
    
    klastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row


    'Find maximum percentage
    For k = 2 To klastrow
       If ws.Cells(k, 11).Value > max Then
        max = ws.Cells(k, 11).Value
        maxticker = ws.Cells(k, 9).Value
       End If

    Next k

    ws.Range("Q2").Value = Format(max, "Percent")
    ws.Range("P2").Value = maxticker
    
    'Find minimum percentage
    For k = 2 To klastrow
       If ws.Cells(k, 11).Value < min Then
        min = ws.Cells(k, 11).Value
        minticker = ws.Cells(k, 9).Value
       End If
    
    Next k
    
    ws.Range("Q3").Value = Format(min, "Percent")
    ws.Range("P3").Value = minticker
    
    'Find maximum volume
    For k = 2 To klastrow
       If ws.Cells(k, 12).Value > max2 Then
        max2 = ws.Cells(k, 12).Value
        maxvolume = ws.Cells(k, 9).Value
       End If
    
    Next k
    
    ws.Range("Q4").Value = max2
    ws.Range("P4").Value = maxvolume

Next ws

End Sub

