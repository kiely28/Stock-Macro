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



Sub CreatePivot_PlantAsRow_OthersAsValues()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim ptCache As PivotCache, pt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim colNumbers As Variant
    Dim headers() As String
    Dim i As Long

    ' Source data sheet
    Set wsData = ThisWorkbook.Sheets("Sheet1")
    
    ' Column numbers: 2 is "Plant" (Row), others are Values
    colNumbers = Array(2, 9, 10, 11, 12, 13, 14, 16, 17, 15, 18)

    ' Convert column numbers to header names
    ReDim headers(LBound(colNumbers) To UBound(colNumbers))
    For i = LBound(colNumbers) To UBound(colNumbers)
        headers(i) = wsData.Cells(1, colNumbers(i)).Value
    Next i

    ' Get last row and last column
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' Delete old PivotOutput if exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PivotOutput").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add new sheet for Pivot Table
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotOutput"

    ' Create pivot cache and table
    Set ptCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)
    Set pt = ptCache.CreatePivotTable(wsPivot.Range("A3"), "CustomPivot")

    ' Set "Plant" as Row Label (first header)
    pt.PivotFields(headers(0)).Orientation = xlRowField
    pt.PivotFields(headers(0)).Position = 1

    ' Add remaining columns as Values (Sum)
    For i = 1 To UBound(headers)
        On Error Resume Next
        pt.AddDataField pt.PivotFields(headers(i)), "Sum of " & headers(i), xlSum
        On Error GoTo 0
    Next i

    ' Format output
    wsPivot.Columns.AutoFit
    MsgBox "Pivot Table created with 'Plant' as Row and custom Values order!", vbInformation
End Sub




---

✅ 1. Put all the macros in a Standard Module

1. Press ALT + F11 to open the VBA editor.


2. In the Project Explorer, right-click on any existing object (like VBAProject (YourWorkbookName.xlsm)).


3. Choose Insert > Module.


4. Paste all the macros (InitButtons, Button1_Click, Button2_Click, Button3_Click) into the module.



📌 Example:

Sub InitButtons()
    With ActiveSheet.Shapes("Button2")
        .OnAction = ""
        .Fill.ForeColor.TintAndShade = 0.8
    End With
    With ActiveSheet.Shapes("Button3")
        .OnAction = ""
        .Fill.ForeColor.TintAndShade = 0.8
    End With
End Sub

Sub Button1_Click()
    MsgBox "Button 1 clicked. Enabling Button 2."
    With ActiveSheet.Shapes("Button2")
        .OnAction = "Button2_Click"
        .Fill.ForeColor.TintAndShade = 0
    End With
End Sub

Sub Button2_Click()
    MsgBox "Button 2 clicked. Enabling Button 3."
    With ActiveSheet.Shapes("Button3")
        .OnAction = "Button3_Click"
        .Fill.ForeColor.TintAndShade = 0
    End With
End Sub

Sub Button3_Click()
    MsgBox "Button 3 clicked. Final action!"
End Sub


---

✅ 2. Assign the Button1 macro to Shape1

In Excel:

1. Right-click your Shape1 (the first "button").


2. Choose Assign Macro.


3. Select Button1_Click.



Repeat for Shape2 and Shape3 (Button2 and Button3), but they won’t work until enabled by code.


---

✅ 3. (Optional) Call InitButtons when opening the sheet

If you want the buttons to reset every time the sheet is activated:

🧩 In the Sheet Module:

1. In the VBA editor, double-click Sheet1 (or your sheet name) under Microsoft Excel Objects.


2. Paste this code:



Private Sub Worksheet_Activate()
    InitButtons
End Sub


---


Sub CreatePivot_PlantAsRow_OthersAsValues()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim ptCache As PivotCache, pt As PivotTable
    Dim dataRange As Range, copyRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim colNumbers As Variant
    Dim headers() As String
    Dim i As Long, dataStartRow As Long, dataEndRow As Long
    Dim pivotTableRange As Range

    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    Dim templatePath As String, savePath As String
    Dim destStartCell As Range, destEndCell As Range
    Dim rowCount As Long, colCount As Long

    ' === Step 1: Setup and Create Pivot ===

    Set wsData = ThisWorkbook.Sheets("Sheet1")
    colNumbers = Array(2, 9, 10, 11, 12, 13, 14, 16, 17, 15, 18)

    ReDim headers(LBound(colNumbers) To UBound(colNumbers))
    For i = LBound(colNumbers) To UBound(colNumbers)
        headers(i) = wsData.Cells(1, colNumbers(i)).Value
    Next i

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PivotOutput").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotOutput"

    Set ptCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)
    Set pt = ptCache.CreatePivotTable(wsPivot.Range("A3"), "CustomPivot")

    pt.PivotFields(headers(0)).Orientation = xlRowField
    pt.PivotFields(headers(0)).Position = 1

    For i = 1 To UBound(headers)
        On Error Resume Next
        pt.AddDataField pt.PivotFields(headers(i)), "Sum of " & headers(i), xlSum
        On Error GoTo 0
    Next i

    wsPivot.Columns.AutoFit
    DoEvents

    ' === Step 2: Copy Pivot Table Data (without header and grand total) ===

    Set pivotTableRange = pt.TableRange1

    If pivotTableRange.Rows.Count > 2 Then
        dataStartRow = pivotTableRange.Row + 1
        dataEndRow = pivotTableRange.Row + pivotTableRange.Rows.Count - 2

        Set copyRange = wsPivot.Range(wsPivot.Cells(dataStartRow, pivotTableRange.Column), _
                                      wsPivot.Cells(dataEndRow, pivotTableRange.Column + pivotTableRange.Columns.Count - 1))

        copyRange.Copy
    Else
        MsgBox "Pivot Table does not contain enough data to copy.", vbExclamation
        Exit Sub
    End If

    ' === Step 3: Open Template and Paste Dynamically from A2 ===

    templatePath = "D:\template.xlsx"
    On Error Resume Next
    Set wbTemplate = Workbooks.Open(templatePath)
    On Error GoTo 0

    If wbTemplate Is Nothing Then
        MsgBox "Template file not found at " & templatePath, vbCritical
        Exit Sub
    End If

    Set wsTemplate = wbTemplate.Sheets(1)

    ' Determine dynamic paste range
    Set destStartCell = wsTemplate.Range("A2")
    rowCount = copyRange.Rows.Count
    colCount = copyRange.Columns.Count
    Set destEndCell = destStartCell.Offset(rowCount - 1, colCount - 1)

    ' Clear and paste into dynamic range
    With wsTemplate.Range(destStartCell, destEndCell)
        .ClearContents
        .PasteSpecial xlPasteValues
    End With

    ' === Step 4: Save As a Copy ===
    savePath = "D:\output\folderresult\pivotdata.xlsx"

    On Error Resume Next
    ' Create directory if it doesn't exist
    If Dir("D:\output\folderresult\", vbDirectory) = "" Then
        MkDir "D:\output"
        MkDir "D:\output\folderresult"
    End If
    On Error GoTo 0

    ' Save As
    Application.DisplayAlerts = False
    wbTemplate.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook ' .xlsx format
    Application.DisplayAlerts = True

    MsgBox "Pivot data copied and saved as 'pivotdata.xlsx' in folderresult!", vbInformation
End Sub




Validation for Difference Posting Check - 070225

Sub CheckPlantAndFilePath()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim plantCount As Long
    Dim i As Long
    Dim missingPath As Boolean
    Dim invalidPath As Boolean
    Dim filePath As String
    Dim fileExt As String
    
    Set ws = ThisWorkbook.Sheets(1) ' Adjust as needed

    ' Find last row in column A (Plant)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Count non-empty plants (excluding header)
    plantCount = WorksheetFunction.CountA(ws.Range("A2:A" & lastRow))

    ' If not exactly 4 plants, validate file paths
    If plantCount <> 4 Then
        missingPath = False
        invalidPath = False
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> "" Then ' has Plant
                filePath = Trim(ws.Cells(i, 2).Value)
                
                If filePath = "" Then
                    ' Missing path
                    missingPath = True
                    ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206) ' Red
                Else
                    ' Check if path is an existing file (not folder)
                    If Dir(filePath, vbNormal) <> "" Then
                        ' Check for Excel file extension
                        fileExt = LCase(Right(filePath, Len(filePath) - InStrRev(filePath, ".")))
                        If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Then
                            ws.Cells(i, 2).Interior.ColorIndex = xlNone ' Valid
                        Else
                            invalidPath = True
                            ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206) ' Not Excel file
                        End If
                    Else
                        invalidPath = True
                        ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206) ' Not found or is a folder
                    End If
                End If
            End If
        Next i
        
        If missingPath Then
            MsgBox "You must enter Material File Path for each Plant (because Plant count ≠ 4).", vbExclamation
        ElseIf invalidPath Then
            MsgBox "One or more paths are not valid Excel files or do not exist.", vbCritical
        Else
            MsgBox "Plant count is not 4, but all Material File Paths are valid Excel files.", vbInformation
        End If
    Else
        ' Exactly 4 plants — clear formatting
        ws.Range("B2:B" & lastRow).Interior.ColorIndex = xlNone
        MsgBox "Exactly 4 Plants — no need to check Material File Path.", vbInformation
    End If
End Sub
