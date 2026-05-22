Attribute VB_Name = "Module1"
Option Explicit

Sub MergeAllFiles()
    Dim fd As FileDialog
    Dim vrtItem As Variant
    Dim sourceWB As Workbook
    Dim targetWS As Worksheet
    Dim sourceWS As Worksheet
    Dim sourceLastRow As Long
    Dim targetLastRow As Long
    Dim totalRows As Long
    Dim fileCounter As Integer
    Dim fileName As String
    Dim rowsCopied As Long
    Dim i As Long
    Dim lastDataRow As Long
    Dim copyRow As Long
    Dim pasteRow As Long
    Dim claimID As String
    Dim newClaimID As String
    Dim fileNum As String
    
    On Error GoTo ErrorHandler
    
    totalRows = 0
    fileCounter = 0
    
    ' Create target worksheet
    Set targetWS = ThisWorkbook.Worksheets.Add
    targetWS.Name = "Merged_Claims"
    
    ' Add headers including Source File column
    With targetWS
        .Range("A1").Value = "Claim ID"
        .Range("B1").Value = "Claim Date"
        .Range("C1").Value = "Insured Name"
        .Range("D1").Value = "Claim Amount"
        .Range("E1").Value = "Status"
        .Range("F1").Value = "Age"
        .Range("G1").Value = "Province"
        .Range("H1").Value = "Source File"
        .Rows("1:1").Font.Bold = True
    End With
    
    ' File selection dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the claim data files to merge"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm", 1
        
        If .Show = -1 Then
            Application.ScreenUpdating = False
            
            For Each vrtItem In .SelectedItems
                fileCounter = fileCounter + 1
                
                ' Open source file
                Set sourceWB = Workbooks.Open(vrtItem)
                Set sourceWS = sourceWB.Worksheets(1)
                
                ' Find last row in source file
                sourceLastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row
                
                If sourceLastRow >= 2 Then
                    
                    ' Find last row in target sheet
                    targetLastRow = targetWS.Cells(targetWS.Rows.Count, "A").End(xlUp).Row + 1
                    If targetLastRow = 2 And targetWS.Range("A2").Value = "" Then
                        targetLastRow = 2
                    End If
                    
                    ' Extract file name and number from full path
                    fileName = Mid(vrtItem, InStrRev(vrtItem, "\") + 1)
                    
                    ' Extract file number (assumes filename contains number like "claim_data_1.xlsx")
                    fileNum = ""
                    For i = 1 To Len(fileName)
                        If IsNumeric(Mid(fileName, i, 1)) Then
                            fileNum = fileNum & Mid(fileName, i, 1)
                        End If
                    Next i
                    
                    If fileNum = "" Then
                        fileNum = CStr(fileCounter)
                    End If
                    
                    ' Copy only rows with valid Claim ID (starts with "CLM")
                    pasteRow = targetLastRow
                    rowsCopied = 0
                    
                    For copyRow = 2 To sourceLastRow
                        claimID = sourceWS.Cells(copyRow, 1).Value
                        
                        ' Check if Claim ID starts with "CLM" (skip TOTAL/AVERAGE rows)
                        If Left(claimID, 3) = "CLM" Then
                            ' Copy all columns from A to G
                            sourceWS.Range("A" & copyRow & ":G" & copyRow).Copy _
                                Destination:=targetWS.Range("A" & pasteRow)
                            
                            ' Modify Claim ID to include file number for uniqueness
                            newClaimID = claimID & "_" & fileNum
                            targetWS.Cells(pasteRow, 1).Value = newClaimID
                            
                            pasteRow = pasteRow + 1
                            rowsCopied = rowsCopied + 1
                        End If
                    Next copyRow
                    
                    ' Fill Source File column (column H)
                    For i = 1 To rowsCopied
                        targetWS.Cells(targetLastRow + i - 1, 8).Value = fileName
                    Next i
                    
                    totalRows = totalRows + rowsCopied
                End If
                
                sourceWB.Close SaveChanges:=False
            Next vrtItem
            
            Application.ScreenUpdating = True
            
            ' Format merged report
            With targetWS
                .Columns("B").NumberFormat = "yyyy-mm-dd"
                .Columns("D").NumberFormat = "$#,##0.00"
                
                ' Add total row
                lastDataRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                .Cells(lastDataRow + 1, 1).Value = "TOTAL"
                .Cells(lastDataRow + 1, 4).Formula = "=SUM(D2:D" & lastDataRow & ")"
                .Rows(lastDataRow + 1).Font.Bold = True
                
                .Cells(lastDataRow + 2, 1).Value = "TOTAL ROWS"
                .Cells(lastDataRow + 2, 4).Value = totalRows
                
                ' Force calculation before auto-fit
                Application.Calculate
                
                ' Auto-fit columns
                .Columns("A:H").AutoFit
            End With
            
            MsgBox "Merge Complete!" & vbNewLine & _
                   "Files processed: " & fileCounter & vbNewLine & _
                   "Total rows: " & totalRows & vbNewLine & _
                   "Summary sheet: " & targetWS.Name, vbInformation
            
        Else
            MsgBox "No files selected.", vbExclamation
        End If
    End With
    
    Set fd = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

