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

- Difference Posting SAP Scripting

  Sub CheckPlantAndSendToSAP()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim plantDict As Object, plantList() As String
    Dim plantVal As String, filePath As String
    Dim sapApp, sapCon, session As Object
    Dim plantCount As Long
    Dim wbMaterial As Workbook
    Dim materialData As Variant
    
    Set ws = ThisWorkbook.Sheets(1)
    Set plantDict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Collect unique plants
    For i = 2 To lastRow
        plantVal = Trim(ws.Cells(i, 1).Value)
        If plantVal <> "" Then
            If Not plantDict.exists(plantVal) Then
                plantDict.Add plantVal, 1
            End If
        End If
    Next i

    plantCount = plantDict.Count

    ' Connect to SAP GUI
    On Error Resume Next
    Set sapApp = GetObject("SAPGUI").GetScriptingEngine
    If sapApp Is Nothing Then
        MsgBox "SAP GUI is not running.", vbCritical
        Exit Sub
    End If
    Set sapCon = sapApp.Children(0)
    Set session = sapCon.Children(0)
    On Error GoTo 0

    ' ✅ SCENARIO 1: Exactly 4 Plants
    If plantCount = 4 Then
        ReDim plantList(0 To plantCount - 1)
        i = 0
        For Each plantVal In plantDict.Keys
            plantList(i) = plantVal
            i = i + 1
        Next

        session.findById("wnd[0]/usr/ctxtPLANT_FIELD").Text = plantList(0)
        session.findById("wnd[0]/usr/btnPLANT_MULTI_BTN").Press
        For i = 1 To 3
            session.findById("wnd[1]/usr/tblPLANT_TABLE/txtPLANT_CELL" & Format(i - 1, "0000")).Text = plantList(i)
        Next i
        session.findById("wnd[1]/tbar[0]/btn[8]").Press ' OK
        session.findById("wnd[0]/tbar[1]/btn[8]").Press ' Execute
        MsgBox "SAP executed for 4 plants.", vbInformation
        Exit Sub
    End If

    ' ❌ SCENARIO 2: Not 4 Plants — process each row
    For i = 2 To lastRow
        plantVal = Trim(ws.Cells(i, 1).Value)
        filePath = Trim(ws.Cells(i, 2).Value)

        If plantVal <> "" And filePath <> "" Then
            ' Enter Plant
            session.findById("wnd[0]/usr/ctxtPLANT_FIELD").Text = plantVal
            
            ' Open Material multiple selection
            session.findById("wnd[0]/usr/btnMATERIAL_MULTI_BTN").Press
            
            ' Open material file
            Set wbMaterial = Workbooks.Open(filePath, ReadOnly:=True)
            With wbMaterial.Sheets(1)
                materialData = .Range("A1", .Cells(.Rows.Count, 1).End(xlUp)).Value
            End With
            
            wbMaterial.Close False
            
            ' Paste materials in SAP clipboard list
            For j = 1 To UBound(materialData, 1)
                session.findById("wnd[1]/usr/tblMATERIAL_TABLE/txtMAT_CELL" & Format(j - 1, "0000")).Text = materialData(j, 1)
            Next j
            
            ' Confirm multiple selection
            session.findById("wnd[1]/tbar[0]/btn[8]").Press ' OK
            session.findById("wnd[0]/tbar[1]/btn[8]").Press ' Execute
        End If
    Next i
    
    MsgBox "SAP executed for all plants with material lists.", vbInformation
End Sub




Code update - 07/04/25


✅ Combined SAP Macro (Scenario 1 + 2 with Clipboard Upload for Materials)

Sub CheckPlantAndSendToSAP_Combined()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim plantDict As Object, plantList() As String
    Dim plantVal As String, filePath As String
    Dim sapApp, sapCon, session As Object
    Dim wbMaterial As Workbook
    Dim materialData As Variant
    Dim clipboardText As String
    Dim plantCount As Long

    ' Set Excel worksheet and get last row
    Set ws = ThisWorkbook.Sheets(1)
    Set plantDict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Collect unique plants into dictionary
    For i = 2 To lastRow
        plantVal = Trim(ws.Cells(i, 1).Value)
        If plantVal <> "" Then
            If Not plantDict.exists(plantVal) Then
                plantDict.Add plantVal, 1
            End If
        End If
    Next i

    plantCount = plantDict.Count

    ' Connect to SAP GUI
    On Error Resume Next
    Set sapApp = GetObject("SAPGUI").GetScriptingEngine
    If sapApp Is Nothing Then
        MsgBox "SAP GUI is not running.", vbCritical
        Exit Sub
    End If
    Set sapCon = sapApp.Children(0)
    Set session = sapCon.Children(0)
    On Error GoTo 0

    ' ✅ SCENARIO 1: If exactly 4 unique plants
    If plantCount = 4 Then
        ' Prepare plant list
        ReDim plantList(0 To plantCount - 1)
        i = 0
        For Each plantVal In plantDict.Keys
            plantList(i) = plantVal
            i = i + 1
        Next

        ' Enter first plant
        session.findById("wnd[0]/usr/ctxtPLANT_FIELD").Text = plantList(0)
        session.findById("wnd[0]/usr/btnPLANT_MULTI_BTN").Press

        ' Enter remaining 3 into multiple selection
        For i = 1 To 3
            session.findById("wnd[1]/usr/tblPLANT_TABLE/txtPLANT_CELL" & Format(i - 1, "0000")).Text = plantList(i)
        Next i

        session.findById("wnd[1]/tbar[0]/btn[8]").Press ' OK
        session.findById("wnd[0]/tbar[1]/btn[8]").Press ' Execute

        MsgBox "SAP executed for 4 unique plants.", vbInformation
        Exit Sub
    End If

    ' ❌ SCENARIO 2: Not exactly 4 — process each row
    For i = 2 To lastRow
        plantVal = Trim(ws.Cells(i, 1).Value)
        filePath = Trim(ws.Cells(i, 2).Value)

        If plantVal <> "" And filePath <> "" Then
            ' Enter plant
            session.findById("wnd[0]/usr/ctxtPLANT_FIELD").Text = plantVal

            ' Open material file
            Set wbMaterial = Workbooks.Open(filePath, ReadOnly:=True)
            With wbMaterial.Sheets(1)
                materialData = .Range("A3", .Cells(.Rows.Count, 1).End(xlUp)).Value
            End With
            wbMaterial.Close False

            ' Convert to clipboard text
            clipboardText = ""
            If IsArray(materialData) Then
                For j = 1 To UBound(materialData, 1)
                    clipboardText = clipboardText & materialData(j, 1) & vbCrLf
                Next j
            Else
                clipboardText = materialData & vbCrLf
            End If

            ' Copy to clipboard
            With CreateObject("htmlfile")
                .ParentWindow.ClipboardData.SetData "text", clipboardText
            End With

            ' Paste materials via clipboard in SAP
            session.findById("wnd[0]/usr/btnMATERIAL_MULTI_BTN").Press
            session.findById("wnd[1]/tbar[0]/btn[24]").Press ' Upload from clipboard
            session.findById("wnd[1]/tbar[0]/btn[8]").Press  ' OK
            session.findById("wnd[0]/tbar[1]/btn[8]").Press  ' Execute
        End If
    Next i

    MsgBox "SAP executed for all plants with material lists.", vbInformation
End Sub


---

Buttons Update 07/04/2025


---

✅ Step 1: Declare 2 Global Flags

In a regular module (e.g., Module1), put this:

Public shape1Clicked As Boolean
Public shape2Clicked As Boolean


---

✅ Step 2: Shape1 Macro

Sub Shape1Macro()
    shape1Clicked = True
    shape2Clicked = False ' Reset shape2 flag in case of re-init
    MsgBox "Shape1 clicked! You can now use Shape2."
    
    ' Put Shape1 action code here
End Sub


---

✅ Step 3: Shape2 Macro

Sub Shape2Macro()
    If shape1Clicked Then
        shape2Clicked = True
        MsgBox "Shape2 running! Shape3 is now available."
        
        ' Put Shape2 action code here
    Else
        MsgBox "Please click Shape1 first before using Shape2."
    End If
End Sub


---

✅ Step 4: Shape3 Macro

Sub Shape3Macro()
    If Not shape1Clicked Then
        MsgBox "Please click Shape1 first before using Shape3."
    ElseIf Not shape2Clicked Then
        MsgBox "Please click Shape2 first before using Shape3."
    Else
        MsgBox "Shape3 running now!"
        
        ' Put Shape3 action code here
    End If
End Sub


---

✅ Step 5 (Optional): Reset on Workbook Open

In ThisWorkbook module:

Private Sub Workbook_Open()
    shape1Clicked = False
    shape2Clicked = False
End Sub


---

Update Scenario 1


Sub SAPScenario1_Loop4RowsFromExcel()
    Dim ws As Worksheet
    Dim i As Long
    Dim sapApp, sapCon, session As Object

    ' Connect to SAP GUI
    On Error Resume Next
    Set sapApp = GetObject("SAPGUI").GetScriptingEngine
    If sapApp Is Nothing Then
        MsgBox "SAP GUI is not running.", vbCritical
        Exit Sub
    End If
    Set sapCon = sapApp.Children(0)
    Set session = sapCon.Children(0)
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets(1)

    ' Optional: check if exactly 4 plant values exist
    If Application.WorksheetFunction.CountA(ws.Range("A2:A5")) <> 4 Then
        MsgBox "Exactly 4 plant values are required in cells A2 to A5.", vbExclamation
        Exit Sub
    End If

    ' Enter first plant directly in SAP
    session.findById("wnd[0]/usr/ctxtPLANT_FIELD").Text = ws.Cells(2, 1).Value

    ' Open SAP multiple selection
    session.findById("wnd[0]/usr/btnPLANT_MULTI_BTN").Press

    ' Enter remaining 3 plants from Excel cells A3 to A5
    For i = 3 To 5
        session.findById("wnd[1]/usr/tblPLANT_TABLE/txtPLANT_CELL" & Format(i - 3, "0000")).Text = ws.Cells(i, 1).Value
    Next i

    ' Confirm and Execute
    session.findById("wnd[1]/tbar[0]/btn[8]").Press  ' OK
    session.findById("wnd[0]/tbar[1]/btn[8]").Press  ' Execute

    MsgBox "SAP executed for plants from A2 to A5.", vbInformation
End Sub


07/05/25 - Inv Results Update

Sub CreateResultsFolder()
    Dim monthYear As String
    Dim mainFolder As String
    Dim subFolder As String
    Dim fullPath As String

    ' Get Month and Year (e.g., "July 2025")
    monthYear = Format(Date, "MMMM YYYY")

    ' Set folder paths
    mainFolder = "D:\Data and Results " & monthYear
    subFolder = mainFolder & "\Inv Results"
    fullPath = subFolder

    ' Check if folder already exists
    If Dir(fullPath, vbDirectory) <> "" Then
        MsgBox "Folder already exists: " & fullPath, vbInformation
    Else
        ' Create main folder if it doesn't exist
        If Dir(mainFolder, vbDirectory) = "" Then
            MkDir mainFolder
        End If

        ' Create subfolder
        MkDir subFolder
        MsgBox "Folder created successfully: " & fullPath, vbInformation
    End If
End Sub




---

✅ Final Macro: Delete Only Excel Files

Sub CreateResultsFolder()
    Dim monthYear As String
    Dim mainFolder As String
    Dim subFolder As String
    Dim fileName As String

    ' Get Month and Year (e.g., "July 2025")
    monthYear = Format(Date, "MMMM YYYY")

    ' Define paths
    mainFolder = "D:\Data and Results " & monthYear
    subFolder = mainFolder & "\Inv Results"

    ' Check if subfolder exists
    If Dir(subFolder, vbDirectory) <> "" Then
        ' Delete only .xlsx and .xls files
        fileName = Dir(subFolder & "\*.xls*") ' Matches .xls and .xlsx
        Do While fileName <> ""
            Kill subFolder & "\" & fileName
            fileName = Dir
        Loop
        MsgBox "Folder already exists. All Excel files were deleted in 'Inv Results'.", vbInformation
    Else
        ' Create folders if they do not exist
        If Dir(mainFolder, vbDirectory) = "" Then MkDir mainFolder
        MkDir subFolder
        MsgBox "Folder created successfully: " & subFolder, vbInformation
    End If
End Sub


---
Update 070725

---

✅ Final Code: Delete Only Excel Files, Show MsgBox for Open Files

Sub CreateResultsFolder()
    Dim monthYear As String
    Dim mainFolder As String
    Dim subFolder As String
    Dim fileName As String
    Dim fullPath As String
    Dim openFiles As String

    ' Get Month and Year (e.g., "July 2025")
    monthYear = Format(Date, "MMMM YYYY")

    ' Define folder paths
    mainFolder = "D:\Data and Results " & monthYear
    subFolder = mainFolder & "\Inv Results"

    ' Check if subfolder exists
    If Dir(subFolder, vbDirectory) <> "" Then
        ' Check and delete Excel files (.xls, .xlsx, .xlsm, etc.)
        fileName = Dir(subFolder & "\*.xls*")
        Do While fileName <> ""
            fullPath = subFolder & "\" & fileName
            If IsFileOpen(fullPath) Then
                openFiles = openFiles & vbCrLf & fileName
            Else
                Kill fullPath
            End If
            fileName = Dir
        Loop

        If openFiles <> "" Then
            MsgBox "Some Excel files could not be deleted because they are open:" & vbCrLf & openFiles, vbExclamation
        Else
            MsgBox "All Excel files deleted in 'Inv Results'.", vbInformation
        End If
    Else
        ' Create folder structure if it doesn't exist
        If Dir(mainFolder, vbDirectory) = "" Then MkDir mainFolder
        MkDir subFolder
        MsgBox "Folder created successfully: " & subFolder, vbInformation
    End If
End Sub

' Function to check if a file is open/locked
Function IsFileOpen(filePath As String) As Boolean
    Dim fileNum As Integer
    Dim errNum As Integer

    On Error Resume Next
    fileNum = FreeFile()
    Open filePath For Binary Access Read Write Lock Read Write As #fileNum
    Close fileNum
    errNum = Err
    On Error GoTo 0

    IsFileOpen = (errNum <> 0)
End Function


---


Dim existingSheet As Worksheet
On Error Resume Next
Set existingSheet = ThisWorkbook.Sheets("PivotOutput")
On Error GoTo 0

If Not existingSheet Is Nothing Then
    Application.DisplayAlerts = False
    existingSheet.Delete
    Application.DisplayAlerts = True
    Set existingSheet = Nothing
End If

Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
wsPivot.Name = "PivotOutput"

Sub CombineFilesCreatePivotAndPasteToTemplate()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook, wbDest As Workbook, wbTemplate As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet, pivotSheet As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim isFirstFile As Boolean
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim plantRange As Range, qtyRange As Range
    Dim copyPlantRange As Range, copyQtyRange As Range
    
    ' Paths
    folderPath = "D:\reports\dpc\"
    Dim templatePath As String
    templatePath = "D:\results\template dpc.xlsx"
    
    ' Create new workbook
    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Sheets(1)
    wsDest.Name = "Combined Data"
    destRow = 1
    isFirstFile = True
    
    ' Combine Excel files
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        Set wbSource = Workbooks.Open(folderPath & fileName)
        Set wsSource = wbSource.Sheets(1)
        
        With wsSource
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            If isFirstFile Then
                .Range("A1:I" & lastRow).Copy wsDest.Range("A" & destRow)
                destRow = destRow + lastRow
                isFirstFile = False
            Else
                .Range("A2:I" & lastRow).Copy wsDest.Range("A" & destRow)
                destRow = destRow + lastRow - 1
            End If
        End With
        
        wbSource.Close False
        fileName = Dir()
    Loop
    
    ' Create Pivot Table
    Set dataRange = wsDest.Range("A1").CurrentRegion
    Set pivotSheet = wbDest.Sheets.Add(After:=wsDest)
    pivotSheet.Name = "Pivot Report"
    
    Set pivotCache = wbDest.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A3"), _
        TableName:="PlantQtyPivot")
    
    With pivotTable
        .PivotFields(1).Orientation = xlRowField  ' Column 1 = Plant
        .PivotFields(9).Orientation = xlDataField ' Column 9 = Qty
        .PivotFields(9).Function = xlSum
        .PivotFields(9).NumberFormat = "#,##0"
    End With
    
    ' Save combined workbook
    Dim combinedPath As String
    combinedPath = folderPath & "combined.xlsx"
    wbDest.SaveAs combinedPath
    
    ' Get Plant and Qty columns from Pivot
    Dim lastPivotRow As Long
    lastPivotRow = pivotSheet.Cells(pivotSheet.Rows.Count, "A").End(xlUp).Row
    
    Set plantRange = pivotSheet.Range("A4:A" & lastPivotRow)
    Set qtyRange = pivotSheet.Range("B4:B" & lastPivotRow)
    
    ' Exclude Grand Total if it exists
    If LCase(plantRange.Cells(plantRange.Rows.Count, 1).Value) = "grand total" Then
        Set copyPlantRange = plantRange.Resize(plantRange.Rows.Count - 1)
        Set copyQtyRange = qtyRange.Resize(qtyRange.Rows.Count - 1)
    Else
        Set copyPlantRange = plantRange
        Set copyQtyRange = qtyRange
    End If
    
    ' Open template file and paste data
    Set wbTemplate = Workbooks.Open(templatePath)
    
    With wbTemplate.Sheets(1)
        .Range("E19").Resize(copyPlantRange.Rows.Count, 1).Value = copyPlantRange.Value
        .Range("G19").Resize(copyQtyRange.Rows.Count, 1).Value = copyQtyRange.Value
    End With
    
    ' Save and close
    wbTemplate.Save
    wbTemplate.Close
    wbDest.Close SaveChanges:=False
    
    MsgBox "✅ Done: Combined, pivoted, and pasted to template."
End Sub


dpc 070925
Sub CombineFilesCreatePivotAndPasteToTemplate()
    Dim folderPath As String, fileName As String
    Dim wbSource As Workbook, wbDest As Workbook, wbTemplate As Workbook, wbPivotInv As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet, pivotSheet As Worksheet, pivotInvSheet As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim isFirstFile As Boolean
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim plantRange As Range, qtyRange As Range
    Dim copyPlantRange As Range, copyQtyRange As Range, invQtyRange As Range, copyInvQtyRange As Range

    ' File Paths
    folderPath = "D:\reports\dpc\"
    Dim templatePath As String: templatePath = "D:\results\template dpc.xlsx"
    Dim pivotInvPath As String: pivotInvPath = "D:\reports\results\pivot inv results.xlsx"
    
    ' Create new workbook for combined data
    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Sheets(1)
    wsDest.Name = "Combined Data"
    destRow = 1
    isFirstFile = True
    
    ' Loop through Excel files and combine data
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        Set wbSource = Workbooks.Open(folderPath & fileName)
        Set wsSource = wbSource.Sheets(1)
        
        With wsSource
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            If isFirstFile Then
                .Range("A1:I" & lastRow).Copy wsDest.Range("A" & destRow)
                destRow = destRow + lastRow
                isFirstFile = False
            Else
                .Range("A2:I" & lastRow).Copy wsDest.Range("A" & destRow)
                destRow = destRow + lastRow - 1
            End If
        End With
        
        wbSource.Close False
        fileName = Dir()
    Loop
    
    ' Create Pivot Table
    Set dataRange = wsDest.Range("A1").CurrentRegion
    Set pivotSheet = wbDest.Sheets.Add(After:=wsDest)
    pivotSheet.Name = "Pivot Report"
    
    Set pivotCache = wbDest.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A3"), _
        TableName:="PlantQtyPivot")
    
    With pivotTable
        .PivotFields(1).Orientation = xlRowField  ' Column 1 = Plant
        .PivotFields(9).Orientation = xlDataField ' Column 9 = Qty
        .PivotFields(9).Function = xlSum
        .PivotFields(9).NumberFormat = "#,##0"
    End With
    
    ' Save combined data
    Dim combinedPath As String: combinedPath = folderPath & "combined.xlsx"
    wbDest.SaveAs combinedPath
    
    ' Get Plant and Qty ranges from pivot
    Dim lastPivotRow As Long
    lastPivotRow = pivotSheet.Cells(pivotSheet.Rows.Count, "A").End(xlUp).Row
    Set plantRange = pivotSheet.Range("A4:A" & lastPivotRow)
    Set qtyRange = pivotSheet.Range("B4:B" & lastPivotRow)
    
    ' Exclude grand total if needed
    If LCase(plantRange.Cells(plantRange.Rows.Count, 1).Value) = "grand total" Then
        Set copyPlantRange = plantRange.Resize(plantRange.Rows.Count - 1)
        Set copyQtyRange = qtyRange.Resize(qtyRange.Rows.Count - 1)
    Else
        Set copyPlantRange = plantRange
        Set copyQtyRange = qtyRange
    End If
    
    ' Open template and paste to E19 (Plant) and G19 (Qty)
    Set wbTemplate = Workbooks.Open(templatePath)
    With wbTemplate.Sheets(1)
        .Range("E19").Resize(copyPlantRange.Rows.Count, 1).Value = copyPlantRange.Value
        .Range("G19").Resize(copyQtyRange.Rows.Count, 1).Value = copyQtyRange.Value
    End With
    
    ' Open pivot inv results file and copy column D
    Set wbPivotInv = Workbooks.Open(pivotInvPath)
    Set pivotInvSheet = wbPivotInv.Sheets("Pivot Result")
    
    Dim lastInvRow As Long
    lastInvRow = pivotInvSheet.Cells(pivotInvSheet.Rows.Count, "D").End(xlUp).Row
    Set invQtyRange = pivotInvSheet.Range("D4:D" & lastInvRow)
    
    ' Exclude grand total if present
    If LCase(invQtyRange.Cells(invQtyRange.Rows.Count, 1).Value) = "grand total" Then
        Set copyInvQtyRange = invQtyRange.Resize(invQtyRange.Rows.Count - 1)
    Else
        Set copyInvQtyRange = invQtyRange
    End If
    
    ' Paste to F19 in template
    wbTemplate.Sheets(1).Range("F19").Resize(copyInvQtyRange.Rows.Count, 1).Value = copyInvQtyRange.Value
    
    ' Save and close all
    wbTemplate.Save
    wbTemplate.Close
    wbPivotInv.Close False
    wbDest.Close SaveChanges:=False
    
    MsgBox "✅ Done: Combined, pivoted, and pasted into template including external qty."
End Sub




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

    ' === Step 1: Prepare Source Sheet and Column Mapping ===
    Set wsData = ThisWorkbook.Sheets("Sheet1")
    colNumbers = Array(2, 9, 10, 11, 12, 13, 14, 16, 17, 15, 18)

    ReDim headers(LBound(colNumbers) To UBound(colNumbers))
    For i = LBound(colNumbers) To UBound(colNumbers)
        headers(i) = wsData.Cells(1, colNumbers(i)).Value
    Next i

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' === Step 2: Delete Existing "PivotOutput" Safely ===
    Dim existingSheet As Worksheet
    On Error Resume Next
    Set existingSheet = ThisWorkbook.Sheets("PivotOutput")
    On Error GoTo 0

    If Not existingSheet Is Nothing Then
        Application.DisplayAlerts = False
        existingSheet.Delete
        Application.DisplayAlerts = True
        Set existingSheet = Nothing
    End If

    ' === Step 3: Create New PivotOutput Sheet ===
    Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPivot.Name = "PivotOutput"

    ' === Step 4: Create Pivot Table ===
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

    ' === Step 5: Copy Pivot Data (excluding header and grand total) ===
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

    ' === Step 6: Open Template File and Paste ===
    templatePath = "D:\template.xlsx"
    On Error Resume Next
    Set wbTemplate = Workbooks.Open(templatePath)
    On Error GoTo 0

    If wbTemplate Is Nothing Then
        MsgBox "Template file not found at " & templatePath, vbCritical
        Exit Sub
    End If

    Set wsTemplate = wbTemplate.Sheets(1)

    Set destStartCell = wsTemplate.Range("A2")
    rowCount = copyRange.Rows.Count
    colCount = copyRange.Columns.Count
    Set destEndCell = destStartCell.Offset(rowCount - 1, colCount - 1)

    With wsTemplate.Range(destStartCell, destEndCell)
        .ClearContents
        .PasteSpecial xlPasteValues
    End With

    ' === Step 7: Save As New File ===
    savePath = "D:\results\reports\pivot inv result.xlsx"

    ' Create folders if missing
    On Error Resume Next
    If Dir("D:\results", vbDirectory) = "" Then MkDir "D:\results"
    If Dir("D:\results\reports", vbDirectory) = "" Then MkDir "D:\results\reports"
    On Error GoTo 0

    Application.DisplayAlerts = False
    wbTemplate.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbTemplate.Close SaveChanges:=False ' ✅ Automatically close the saved file
    Application.DisplayAlerts = True

    MsgBox "Pivot data pasted and saved as 'pivot inv result.xlsx' in D:\results\reports. File has been closed.", vbInformation
End Sub



07/10/25 - SAP Check Error
---

--- Full Example

Sub SAP_ExportCheck()

    Dim session As Object
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0).Children(0)

    ' Perform your transaction steps here...
    ' For example, session.findById("...").press

    ' Wait for SAP to load response (you can add delay or wait loop)
    Application.Wait Now + TimeValue("0:00:02")

    ' Check SAP status bar
    Dim statusText As String
    statusText = session.findById("wnd[0]/sbar").Text

    If InStr(statusText, "No data") > 0 Or _
       InStr(statusText, "not found") > 0 Or _
       InStr(statusText, "Error in ABAP") > 0 Or _
       session.findById("wnd[0]/sbar").MessageType = "E" Then

        MsgBox "❌ SAP Error: " & statusText, vbCritical, "SAP Message"
        Exit Sub
    End If

    ' Otherwise continue exporting to Excel
    MsgBox "✅ Data loaded successfully!", vbInformation

End Sub


---
07-10-25 SaveAs Close Workbook

---

✅ Updated Macro: Save and Close Workbook

Sub SaveAndCloseWorkbook()

    Dim savePath As String
    Dim fileName As String
    Dim fullPath As String

    ' Set the path and file name
    savePath = "D:\results\reports\"           ' Your desired folder
    fileName = "Report_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    fullPath = savePath & fileName

    ' Check if file already exists (optional safety)
    If Dir(fullPath) <> "" Then
        MsgBox "File already exists: " & fullPath, vbExclamation
        Exit Sub
    End If

    ' Save the workbook
    ThisWorkbook.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook ' .xlsx format

    ' Confirm save
    MsgBox "Workbook saved as: " & fullPath, vbInformation

    ' Close the workbook
    ThisWorkbook.Close SaveChanges:=False ' Already saved, no need to save again

End Sub


---

🔍 Key Parts Explained

Code Line	Purpose

ThisWorkbook.SaveAs ...	Saves the current workbook to the specified path
ThisWorkbook.Close SaveChanges:=False	Closes the workbook without asking to save again
Dir(fullPath) check	Prevents overwriting existing files (optional)
MsgBox	Confirms success before closing



---

🛡️ Tip:

If this workbook runs from a master macro file, and you don't want to close it, use:

ActiveWorkbook.Close SaveChanges:=False

instead of ThisWorkbook.




---

✅ Updated Part of the Code

This part saves the wsData (Sheet1) into:

📁 D:\reports\results\pivot summary inv results.xlsx

' === Step 7: Save wsData (Sheet1) to D:\reports\results\pivot summary inv results.xlsx ===
summaryPath = "D:\reports\results\pivot summary inv results.xlsx"

' Ensure folder exists
On Error Resume Next
If Dir("D:\reports", vbDirectory) = "" Then MkDir "D:\reports"
If Dir("D:\reports\results", vbDirectory) = "" Then MkDir "D:\reports\results"
On Error GoTo 0

' Create new workbook with copy of wsData
wsData.Copy
Set wbSummary = ActiveWorkbook
Application.DisplayAlerts = False
wbSummary.SaveAs Filename:=summaryPath, FileFormat:=xlOpenXMLWorkbook
wbSummary.Close SaveChanges:=False
Application.DisplayAlerts = True


---


Update 07/10/25

---

🧾 Reusable Subroutine

Sub AddOrCleanSubfolder()
    Dim mainFolder As String
    Dim subFolderName As String
    Dim subFolderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim openFiles As String

    ' === Configuration ===
    mainFolder = "D:\Data and Results July 2025" ' Change this to your existing main folder path
    subFolderName = "Summary Files"             ' Subfolder name you want to add
    subFolderPath = mainFolder & "\" & subFolderName

    ' === Check if subfolder exists ===
    If Dir(subFolderPath, vbDirectory) <> "" Then
        ' Subfolder exists — delete Excel files only
        fileName = Dir(subFolderPath & "\*.xls*")
        Do While fileName <> ""
            fullPath = subFolderPath & "\" & fileName
            If IsFileOpen(fullPath) Then
                openFiles = openFiles & vbCrLf & fileName
            Else
                Kill fullPath
            End If
            fileName = Dir
        Loop

        If openFiles <> "" Then
            MsgBox "Some Excel files could not be deleted because they are open:" & vbCrLf & openFiles, vbExclamation
        Else
            MsgBox "All Excel files in '" & subFolderName & "' were deleted.", vbInformation
        End If
    Else
        ' Subfolder does not exist — create it
        MkDir subFolderPath
        MsgBox "Subfolder created: " & subFolderPath, vbInformation
    End If
End Sub


---

🔁 Required Helper Function

Make sure you include this function in a module (if you haven't yet):

Function IsFileOpen(filePath As String) As Boolean
    Dim fileNum As Integer
    Dim errNum As Integer

    On Error Resume Next
    fileNum = FreeFile()
    Open filePath For Binary Access Read Write Lock Read Write As #fileNum
    Close fileNum
    errNum = Err
    On Error GoTo 0

    IsFileOpen = (errNum <> 0)
End Function


---

✅ Customize Easily

Change the following lines to target any folder/subfolder:

mainFolder = "D:\Data and Results July 2025"
subFolderName = "Summary Files"


---


