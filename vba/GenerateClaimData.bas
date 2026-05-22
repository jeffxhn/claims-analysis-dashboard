Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateClaimData()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Use Timer as random seed for better randomness
    Randomize Timer
    
    ' Create new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Claims_Data_" & Format(Now, "yyyymmdd_hhmmss")
    
    ' Add headers
    With ws
        .Range("A1").Value = "Claim ID"
        .Range("B1").Value = "Claim Date"
        .Range("C1").Value = "Insured Name"
        .Range("D1").Value = "Claim Amount"
        .Range("E1").Value = "Status"
        .Range("F1").Value = "Age"
        .Range("G1").Value = "Province"
        .Rows("1:1").Font.Bold = True
    End With
    
    ' Generate 100 rows of random data
    For i = 2 To 101
        ' Claim ID: CLM001, CLM002, ...
        ws.Cells(i, 1).Value = "CLM" & Format(i - 1, "000")
        
        ' Claim Date: random date between Jan 1, 2024 and Dec 31, 2024
        ws.Cells(i, 2).Value = DateSerial(2024, 1, 1) + Int(Rnd() * 365)
        
        ' Insured Name: random from A, B, C, D
        ws.Cells(i, 3).Value = Choose(Int(Rnd() * 4) + 1, "A", "B", "C", "D")
        
        ' Claim Amount: random between 1000 and 100000
        ws.Cells(i, 4).Value = Int(Rnd() * 99001) + 1000
        
        ' Status: random from Approved, Pending, Rejected
        ws.Cells(i, 5).Value = Choose(Int(Rnd() * 3) + 1, "Approved", "Pending", "Rejected")
        
        ' Age: random between 18 and 70
        ws.Cells(i, 6).Value = Int(Rnd() * 53) + 18
        
        ' Province: random from ON, BC, AB, QC, MB
        ws.Cells(i, 7).Value = Choose(Int(Rnd() * 5) + 1, "ON", "BC", "AB", "QC", "MB")
    Next i
    
    ' Format columns
    ws.Columns("A").NumberFormat = "@"
    ws.Columns("B").NumberFormat = "yyyy-mm-dd"
    ws.Columns("D").NumberFormat = "$#,##0.00"
    ws.Columns("A:G").AutoFit
    
    MsgBox "Generated 100 rows of claim data in sheet: " & ws.Name, vbInformation
End Sub
