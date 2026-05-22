Attribute VB_Name = "Module2"
Option Explicit

Sub FormatOneReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim amountCol As Long
    Dim totalRow As Long
    Dim i As Long
    Dim threshold As Double
    
    Set ws = ActiveSheet
    amountCol = 4
    threshold = 50000
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found.", vbExclamation
        Exit Sub
    End If
    
    ' Format header row
    With ws.Rows("1:1")
        .Font.Bold = True
        .Interior.Color = RGB(211, 211, 211)
    End With
    
    ' Format amount column as currency
    ws.Columns(amountCol).NumberFormat = "$#,##0.00"
    
    ' Highlight high claim amounts (red and bold)
    For i = 2 To lastRow
        If ws.Cells(i, amountCol).Value > threshold Then
            ws.Cells(i, amountCol).Font.Color = RGB(255, 0, 0)
            ws.Cells(i, amountCol).Font.Bold = True
        End If
    Next i
    
    ' Add summary rows
    totalRow = lastRow + 1
    ws.Cells(totalRow, 1).Value = "TOTAL"
    ws.Cells(totalRow, amountCol).Formula = "=SUM(D2:D" & lastRow & ")"
    
    ws.Cells(totalRow + 1, 1).Value = "AVERAGE"
    ws.Cells(totalRow + 1, amountCol).Formula = "=AVERAGE(D2:D" & lastRow & ")"
    
    ws.Rows(totalRow & ":" & totalRow + 1).Font.Bold = True
    
    ' Auto-fit columns
    ws.Columns("A:G").AutoFit
    
    MsgBox "Formatting complete. Processed " & (lastRow - 1) & " rows.", vbInformation
End Sub
