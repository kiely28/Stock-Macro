# Stock-Macro
Macro


Sub CleanFormatSortAndNumber_WithIncharge()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim lastRow As Long, i As Long
    Dim allZero As Boolean: allZero = True

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim matrixWB As Workbook
    Dim matrixWS As Worksheet
    Dim matrixPath As String
    Dim matrixRow As Long
    Dim locKey As String, inchargeVal As String
    Dim mode As String

    Application.ScreenUpdating = False

    ' === Step 1: Remove "Total" Rows ===
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If Trim(UCase(ws.Cells(i, "A").Value)) = "TOTAL" Then
            ws.Rows(i).Delete
        End If
    Next i

    ' === Step 2: Incharge Mapping ===
    ws.Columns("E").Insert Shift:=xlToRight
    ws.Cells(1, 5).Value = "Incharge"

    matrixPath = "D:\Path\To\matrix.xlsx" ' Update as needed
    On Error Resume Next
    Set matrixWB = Workbooks.Open(matrixPath, ReadOnly:=True)
    On Error GoTo 0

    If matrixWB Is Nothing Then
        MsgBox "Matrix file not found!", vbCritical
        Exit Sub
    End If

    Set matrixWS = matrixWB.Sheets(1)
    matrixRow = matrixWS.Cells(matrixWS.Rows.Count, "A").End(xlUp).Row

    For i = 2 To matrixRow
        locKey = Trim(matrixWS.Cells(i, 1).Value)
        inchargeVal = Trim(matrixWS.Cells(i, 2).Value)
        If Not dict.exists(LCase(locKey)) Then dict.Add LCase(locKey), inchargeVal
    Next i
    matrixWB.Close SaveChanges:=False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        locKey = Trim(ws.Cells(i, 4).Value)
        If dict.exists(LCase(locKey)) Then
            ws.Cells(i, 5).Value = dict(LCase(locKey))
        Else
            ws.Cells(i, 5).Value = "N/A"
            ws.Cells(i, 5).Interior.Color = RGB(255, 255, 0)
        End If
    Next i

    ' === Step 3: Determine if block qty present ===
    allZero = True
    For i = 2 To lastRow
        If ws.Cells(i, 24).Value <> 0 Or _
           ws.Cells(i, 25).Value <> 0 Or _
           ws.Cells(i, 26).Value <> 0 Or _
           ws.Cells(i, 27).Value <> 0 Then
            allZero = False
            Exit For
        End If
    Next i

    ' === Step 4: Delete columns and format ===
    If allZero Then
        mode = "NoBlock"
        ws.Range("W:AI").Delete
        ws.Columns("S").Delete
        ws.Range("G:J").Delete
    Else
        mode = "WithBlock"
        ws.Columns("W").Delete
        ws.Columns("S").Delete
        ws.Range("G:J").Delete
    End If

    ' === Step 5: Apply custom number formats ===
    Dim col As Range
    For Each col In ws.Range("H1:AC1")
        Select Case col.Column
            Case 8, 10, 12, 13, 14, 18, 20, 22, 23, 24 ' Qty Columns
                ws.Columns(col.Column).NumberFormat = "#,##0"
            Case 7, 9, 11, 15, 16, 17, 19, 22, 27, 28   ' Amt Columns
                ws.Columns(col.Column).NumberFormat = "#,##0.00"
        End Select
    Next col

    ' === Step 6: Sort by Column O ===
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("O2:O" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:Z" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' === Step 7: Add "No" column ===
    ws.Cells(1, "A").Value = "No"
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        ws.Cells(i, "A").Value = i - 1
    Next i

    ' === Step 8: Insert Totals & Percentage ===
    Call InsertTotalRow(ws, mode)

    Application.ScreenUpdating = True
    MsgBox "✅ Macro completed successfully!", vbInformation
End Sub

' ===========================
' === Subroutine: Totals ===
' ===========================
Sub InsertTotalRow(ws As Worksheet, mode As String)
    Dim lastRow As Long, i As Long
    Dim startCol As Integer, endCol As Integer
    Dim totalRow As Long, blockTotalRow As Long, grandTotalRow As Long
    Dim percentRow As Long

    If mode = "NoBlock" Then
        startCol = 8: endCol = 17 ' H:Q
    ElseIf mode = "WithBlock" Then
        startCol = 8: endCol = 21 ' H:U
    Else
        MsgBox "Invalid mode passed", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ws.Rows(lastRow + 1).Insert

    ' Total Row
    totalRow = lastRow + 2
    For i = startCol To endCol
        ws.Cells(totalRow, i).Formula = "=SUM(" & ws.Cells(2, i).Address & ":" & ws.Cells(lastRow, i).Address & ")"
    Next i
    With ws.Cells(totalRow, 7)
        .Value = "Total"
        .Font.Bold = True
    End With

    If mode = "WithBlock" Then
        ' Block Total
        blockTotalRow = totalRow + 1
        For i = 0 To 3
            ws.Cells(blockTotalRow, 8 + i).Formula = "=SUM(" & ws.Cells(2, 18 + i).Address & ":" & ws.Cells(lastRow, 18 + i).Address & ")"
        Next i
        With ws.Cells(blockTotalRow, 7)
            .Value = "Block Total"
            .Font.Bold = True
        End With

        ' Grand Total
        grandTotalRow = blockTotalRow + 1
        For i = 8 To 11
            ws.Cells(grandTotalRow, i).Formula = "=" & ws.Cells(totalRow, i).Address & "+" & ws.Cells(blockTotalRow, i).Address
        Next i
        With ws.Cells(grandTotalRow, 7)
            .Value = "Grand Total"
            .Font.Bold = True
        End With

        ' Percentage (P / I from Grand Total)
        percentRow = grandTotalRow + 2
        ws.Rows(percentRow).Insert
        ws.Rows(percentRow + 1).Insert
        With ws.Cells(percentRow + 1, 16) ' Column P
            .Formula = "=" & ws.Cells(grandTotalRow, 16).Address & "/" & ws.Cells(grandTotalRow, 9).Address
            .NumberFormat = "0.00%"
            .Interior.Color = RGB(255, 255, 0)
            .Font.Bold = True
        End With
        With ws.Cells(percentRow + 1, 7)
            .Value = "Conversion %"
            .Font.Bold = True
        End With

    ElseIf mode = "NoBlock" Then
        ' Percentage (P / I from Total)
        percentRow = totalRow + 3
        ws.Rows(percentRow).Insert
        ws.Rows(percentRow + 1).Insert
        With ws.Cells(percentRow + 1, 16)
            .Formula = "=" & ws.Cells(totalRow, 16).Address & "/" & ws.Cells(totalRow, 9).Address
            .NumberFormat = "0.00%"
            .Interior.Color = RGB(255, 255, 0)
            .Font.Bold = True
        End With
        With ws.Cells(percentRow + 1, 7)
            .Value = "Conversion %"
            .Font.Bold = True
        End With
    End If
End Sub


' === Step 9: Final Formatting ===
ws.Cells.Font.Name = "Calibri"
ws.Cells.Font.Size = 10
ws.Cells.Interior.Color = RGB(255, 255, 255) ' Set background to white
ws.Rows(1).Font.Bold = True ' Bold the header row
ws.Cells.EntireColumn.AutoFit ' Auto-fit column widths
ws.Activate
ActiveWindow.Zoom = 70 ' Set zoom level to 70%

' === Step 10: Save Workbook ===
ThisWorkbook.Save
