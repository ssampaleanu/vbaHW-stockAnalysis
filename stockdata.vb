Sub stockdata()
'   label output table
Range("i1").Value = "Ticker Symbol"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume (Millions)"
Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker Symbol"
Range("Q1").Value = "Value"

'   declare variables
Dim openval As Double
Dim closeval As Double
Dim totalchange As Double
Dim totalvol As Double

Dim greatestinc As Variant
Dim inctick As Variant

Dim greatestdec As Variant
Dim dectick As Variant

Dim greatestvol As Variant
Dim voltick As Variant

Dim i As Long
Dim j As Long
    j = 2
Dim growthSumm As Range
Dim volSumm As Range

For i = 2 To 1000000
    totalvol = totalvol + (Cells(i, 7).Value / 1000000)

        '   first, check if new ticker val
        '   store opening value of ticker val
    If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
        openval = Cells(i, 3).Value

    ElseIf (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        closeval = Cells(i, 6).Value
        totalchange = closeval - openval
            
            ' print output to the right of table
            Cells(j, 9).Value = Cells(i, 1).Value
            Cells(j, 13).Value = Cells(i, 1).Value
            Cells(j, 10).Value = totalchange
                Cells(j, 10).NumberFormat = "$0.00"
                If Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                Else
                    Cells(j, 10).Interior.ColorIndex = 4
                End If
            Cells(j, 11).Value = totalchange / openval
                On Error Resume Next
                Cells(j, 11).NumberFormat = "0%"
            Cells(j, 12).Value = totalvol
                Cells(j, 12).NumberFormat = "$0.0"
            j = j + 1
            totalvol = 0
    ElseIf IsEmpty(Cells(i, 1).Value) Then Exit For
        
    End If
Next i
    
With ActiveSheet
    Set growthSumm = ActiveSheet.Columns("k:m")
    Set volSumm = ActiveSheet.Columns("L:M")
End With
        '   Hard difficulty: print the greatest increase, decrease, and volume to the left of table
        '   use vlookup on active sheet to show ticker symbol for each value
    greatestinc = Application.WorksheetFunction.Max(Columns(11))
        Range("Q2").Value = greatestinc
        inctick = Application.WorksheetFunction.VLookup(greatestinc, growthSumm, 3, 0)
        Range("P2").Value = inctick
    greatestdec = Application.WorksheetFunction.Min(Columns(11))
        Range("Q3").Value = greatestdec
        dectick = Application.WorksheetFunction.VLookup(greatestdec, growthSumm, 3, 0)
        Range("P3").Value = dectick
    Range("q2:q3").NumberFormat = "0%"
    
    grtotalvol = Application.WorksheetFunction.Max(Columns(12))
        Range("Q4").Value = grtotalvol
        Range("Q4").NumberFormat = "$0.0"
        voltick = Application.WorksheetFunction.VLookup(grtotalvol, volSumm, 2, 0)
        Range("P4").Value = voltick
        

End Sub