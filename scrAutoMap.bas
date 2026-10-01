Attribute VB_Name = "scrAutoMap"
Option Explicit

' CONFIG
Private Const CONFIG_HEADER_ROW As Long = 1
Private Const CONFIG_START_ROW As Long = 2
Private Const CONFIG_BLOCKS_PER_ROW As Long = 4
Private Const CONFIG_OUTPUT_START_ROW As Long = 1
Private Const CONFIG_OUTPUT_START_COLUMN As Long = 2
Private Const CONFIG_BLOCK_COLUMN_GAP As Long = 1
Private Const CONFIG_BLOCK_ROW_GAP As Long = 1
Private Const CONFIG_BLOCK_ROW_GAP_HEIGHT As Double = 15
Private Const CONFIG_BLOCK_COLUMN_GAP_WIDTH As Double = 2
Private Const CONFIG_OUTPUT_PREFIX As String = "CableMap_"
Private Const CONFIG_OUTPUT_EXTENSION As String = ".xlsx"
Private Const CONFIG_FIRST_SHEET_NAME As String = "#Reference"
Private Const CONFIG_SKIP_EMPTY_VALUES As Boolean = True
Private Const CONFIG_DRAW_BORDERS As Boolean = True
Private Const CONFIG_AUTOFIT_COLUMNS As Boolean = True

Public AC_HeaderList As Variant
Public AC_ListColumnIndex() As Long
Public AC_RemoveColumns() As Long
Public AC_RemoveCount As Long
Public AC_GroupIndex As Long
Public AC_Cancelled As Boolean
Public AC_SelectGroupColumn As Boolean

Public Sub ProcessLocationSheets()
    Dim wbSource As Workbook, wbOutput As Workbook
    Dim wsSource As Worksheet, wsWork As Worksheet, wsOutput As Worksheet
    Dim wsFirstSheet As Worksheet, wsExtraSheet As Worksheet
    Dim wsCopiedExtraSheet As Worksheet
    Dim selectedRange As Range
    Dim dataRange As Range
    Dim sourceData As Variant, groups As Object, key As Variant
    Dim removeFlag() As Boolean, colWidths() As Double
    Dim lastRow As Long, lastColumn As Long, rowIndex As Long
    Dim columnIndex As Long, outputColumn As Long, groupColumnIndex As Long
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean, oldCalculation As XlCalculation
    Dim outputPath As String, targetName As String, savedError As String
    Dim answer As VbMsgBoxResult, outputSheetCount As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set wsSource = ActiveSheet
    Set wbSource = wsSource.Parent
    On Error Resume Next
    Set selectedRange = Application.InputBox(Prompt:="Select the source table range.", Title:="AutoCable", Type:=8)
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Sub
    If Not selectedRange.Parent Is wsSource Then
        MsgBox "Select a range on the active source sheet.", vbExclamation, "AutoCable"
        Exit Sub
    End If
    If selectedRange.Areas.Count > 1 Then
        MsgBox "Select one continuous table range.", vbExclamation, "AutoCable"
        Exit Sub
    End If
    If Len(wbSource.Path) = 0 Then
        MsgBox "Save the source workbook before running this macro.", vbExclamation, "AutoCable"
        Exit Sub
    End If

    If Len(Trim$(CONFIG_FIRST_SHEET_NAME)) > 0 Then
        On Error Resume Next
        Set wsExtraSheet = wbSource.Worksheets(CONFIG_FIRST_SHEET_NAME)
        On Error GoTo 0
        If wsExtraSheet Is Nothing Then
            answer = MsgBox("Sheet '" & CONFIG_FIRST_SHEET_NAME & "' was not found." & vbCrLf & _
                            "Yes = Continue without it" & vbCrLf & "No = Stop", _
                            vbQuestion + vbYesNo + vbDefaultButton2, "AutoCable")
            If answer <> vbYes Then Exit Sub
        End If
    End If

    outputPath = wbSource.Path & Application.PathSeparator & _
                 CONFIG_OUTPUT_PREFIX & Format$(Date, "yymmdd") & CONFIG_OUTPUT_EXTENSION
    If Len(Dir$(outputPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then
        answer = MsgBox("Output file exists. Yes = Override; No = Stop." & vbCrLf & outputPath, _
                        vbQuestion + vbYesNo + vbDefaultButton2, "AutoCable")
        If answer <> vbYes Then Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Set wsWork = wbSource.Worksheets.Add(After:=wsSource)
    selectedRange.Copy Destination:=wsWork.Cells(1, 1)
    lastRow = selectedRange.Rows.Count
    lastColumn = selectedRange.Columns.Count
    ReDim colWidths(1 To lastColumn)
    For columnIndex = 1 To lastColumn
        colWidths(columnIndex) = selectedRange.Columns(columnIndex).ColumnWidth
    Next columnIndex

    Set dataRange = wsWork.Cells(1, 1).Resize(lastRow, lastColumn)
    dataRange.Value2 = dataRange.Value2
    ReDim AC_HeaderList(1 To lastColumn)
    For rowIndex = 1 To lastColumn
        AC_HeaderList(rowIndex) = wsWork.Cells(CONFIG_HEADER_ROW, rowIndex).Text
        If Len(AC_HeaderList(rowIndex)) = 0 Then AC_HeaderList(rowIndex) = "(blank header)"
    Next rowIndex

    AC_GroupIndex = 0
    AC_SelectGroupColumn = True
    AC_Cancelled = True
    Application.ScreenUpdating = oldScreenUpdating
    frmAutoMap.Show vbModal
    Unload frmAutoMap
    If AC_Cancelled Then GoTo Cleanup
    If AC_GroupIndex < 1 Or AC_GroupIndex > lastColumn Then Err.Raise vbObjectError + 3, , "Invalid grouping column."

    AC_SelectGroupColumn = False
    AC_Cancelled = True
    AC_RemoveCount = 0
    Erase AC_RemoveColumns
    Application.ScreenUpdating = oldScreenUpdating
    frmAutoMap.Show vbModal
    Unload frmAutoMap
    If AC_Cancelled Then GoTo Cleanup

    Application.ScreenUpdating = False
    ReDim removeFlag(1 To lastColumn)
    For rowIndex = 1 To AC_RemoveCount
        removeFlag(AC_RemoveColumns(rowIndex)) = True
    Next rowIndex
    For rowIndex = AC_RemoveCount To 1 Step -1
        wsWork.Columns(AC_RemoveColumns(rowIndex)).Delete
    Next rowIndex
    groupColumnIndex = AC_GroupIndex
    For columnIndex = 1 To AC_GroupIndex
        If removeFlag(columnIndex) Then groupColumnIndex = groupColumnIndex - 1
    Next columnIndex
    outputColumn = 0
    For columnIndex = 1 To lastColumn
        If Not removeFlag(columnIndex) Then
            outputColumn = outputColumn + 1
            wsWork.Columns(outputColumn).ColumnWidth = colWidths(columnIndex)
        End If
    Next columnIndex
    lastColumn = lastColumn - AC_RemoveCount
    sourceData = wsWork.Cells(1, 1).Resize(lastRow, lastColumn).Value2

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    For rowIndex = CONFIG_START_ROW To UBound(sourceData, 1)
        Dim groupName As String
        groupName = CellText(sourceData(rowIndex, groupColumnIndex))
        If Len(groupName) > 0 Then
            If Not groups.Exists(groupName) Then groups.Add groupName, New Collection
            groups(groupName).Add rowIndex
        End If
    Next rowIndex
    If groups.Count = 0 Then Err.Raise vbObjectError + 2, , "No group values were found in the selected grouping column."

    Set wbOutput = Workbooks.Add(xlWBATWorksheet)
    Set wsFirstSheet = wbOutput.Worksheets(1)
    If Not wsExtraSheet Is Nothing Then
        wsExtraSheet.Copy Before:=wbOutput.Worksheets(1)
        Set wsCopiedExtraSheet = wbOutput.Worksheets(1)
        wsCopiedExtraSheet.PageSetup.PrintArea = wsCopiedExtraSheet.UsedRange.Address
        ApplyPrintSetup wsCopiedExtraSheet
    End If

    outputSheetCount = 0
    For Each key In groups.Keys
        If outputSheetCount = 0 Then
            targetName = UniqueSheetName(wbOutput, SafeSheetName(CStr(key)), wsFirstSheet)
            Set wsOutput = wsFirstSheet
        Else
            targetName = UniqueSheetName(wbOutput, SafeSheetName(CStr(key)), Nothing)
            Set wsOutput = wbOutput.Worksheets.Add(After:=wbOutput.Worksheets(wbOutput.Worksheets.Count))
        End If
        wsOutput.Name = targetName
        WriteGroupOutput wsOutput, wsWork, sourceData, groups(key), lastColumn
        ApplyPrintSetup wsOutput
        outputSheetCount = outputSheetCount + 1
    Next key

    wbOutput.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    MsgBox "Output file created:" & vbCrLf & outputPath, vbInformation, "AutoCable"
    GoTo Cleanup

ErrorHandler:
    savedError = Err.Description
    MsgBox "Processing failed: " & savedError, vbCritical, "AutoCable"

Cleanup:
    On Error Resume Next
    If Not wbOutput Is Nothing Then
        If Len(wbOutput.Path) = 0 Then wbOutput.Close SaveChanges:=False
    End If
    If Not wsWork Is Nothing Then wsWork.Delete
    On Error GoTo 0
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
End Sub

Private Sub WriteGroupOutput(ByVal wsOutput As Worksheet, ByVal wsSource As Worksheet, ByRef sourceData As Variant, _
                             ByVal rowList As Collection, ByVal lastColumn As Long)
    Dim recordCount As Long, blockRows As Long, maxOutputRows As Long
    Dim outputColumns As Long, outputData() As Variant
    Dim recordIndex As Long, rowIndex As Long, columnIndex As Long
    Dim blockRow As Long, blockPosition As Long, itemCount As Long
    Dim blockHeight As Long, rowOffset As Long, columnOffset As Long
    Dim valueIndex As Long, gapIndex As Long
    Dim headerCell As Range, outputCell As Range
    Dim headerFillPattern() As Variant, headerFillColor() As Variant, headerFontColor() As Variant

    recordCount = rowList.Count
    blockRows = (recordCount + CONFIG_BLOCKS_PER_ROW - 1) \ CONFIG_BLOCKS_PER_ROW
    outputColumns = CONFIG_OUTPUT_START_COLUMN + (CONFIG_BLOCKS_PER_ROW - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP) + 1

    For blockRow = 0 To blockRows - 1
        blockHeight = 0
        For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW
            recordIndex = blockRow * CONFIG_BLOCKS_PER_ROW + blockPosition
            If recordIndex <= recordCount Then
                rowIndex = CLng(rowList(recordIndex))
                itemCount = CountOutputItems(sourceData, rowIndex, lastColumn)
                If itemCount > blockHeight Then blockHeight = itemCount
            End If
        Next blockPosition
        maxOutputRows = maxOutputRows + blockHeight
        If blockRow < blockRows - 1 Then maxOutputRows = maxOutputRows + CONFIG_BLOCK_ROW_GAP
    Next blockRow

    If maxOutputRows = 0 Then Exit Sub
    ReDim outputData(1 To maxOutputRows, 1 To outputColumns)
    ReDim headerFillPattern(1 To lastColumn)
    ReDim headerFillColor(1 To lastColumn)
    ReDim headerFontColor(1 To lastColumn)

    For columnIndex = 1 To lastColumn
        Set headerCell = wsSource.Cells(CONFIG_HEADER_ROW, columnIndex)
        headerFillPattern(columnIndex) = headerCell.Interior.Pattern
        headerFillColor(columnIndex) = headerCell.Interior.Color
        headerFontColor(columnIndex) = headerCell.Font.Color
    Next columnIndex

    rowOffset = 1
    For blockRow = 0 To blockRows - 1
        blockHeight = 0
        For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW
            recordIndex = blockRow * CONFIG_BLOCKS_PER_ROW + blockPosition
            If recordIndex <= recordCount Then
                rowIndex = CLng(rowList(recordIndex))
                itemCount = CountOutputItems(sourceData, rowIndex, lastColumn)
                If itemCount > blockHeight Then blockHeight = itemCount
                columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP)
                valueIndex = 0

                For columnIndex = 1 To lastColumn
                    If ShouldOutput(sourceData(rowIndex, columnIndex)) Then
                        valueIndex = valueIndex + 1
                        outputData(rowOffset + valueIndex - 1, columnOffset) = sourceData(CONFIG_HEADER_ROW, columnIndex)
                        outputData(rowOffset + valueIndex - 1, columnOffset + 1) = sourceData(rowIndex, columnIndex)
                    End If
                Next columnIndex
            End If
        Next blockPosition

        If CONFIG_DRAW_BORDERS And blockHeight > 0 Then
            For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW
                recordIndex = blockRow * CONFIG_BLOCKS_PER_ROW + blockPosition
                If recordIndex <= recordCount Then
                    rowIndex = CLng(rowList(recordIndex))
                    itemCount = CountOutputItems(sourceData, rowIndex, lastColumn)
                    If itemCount > 0 Then
                        columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP)
                        With wsOutput.Range(wsOutput.Cells(CONFIG_OUTPUT_START_ROW + rowOffset - 1, columnOffset), _
                                            wsOutput.Cells(CONFIG_OUTPUT_START_ROW + rowOffset + itemCount - 2, columnOffset + 1))
                            .Borders.LineStyle = xlContinuous
                        End With
                    End If
                End If
            Next blockPosition
        End If
        rowOffset = rowOffset + blockHeight + CONFIG_BLOCK_ROW_GAP
    Next blockRow

    wsOutput.Cells(CONFIG_OUTPUT_START_ROW, 1).Resize(maxOutputRows, outputColumns).value = outputData

    rowOffset = 1
    For blockRow = 0 To blockRows - 1
        blockHeight = 0
        For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW
            recordIndex = blockRow * CONFIG_BLOCKS_PER_ROW + blockPosition
            If recordIndex <= recordCount Then
                rowIndex = CLng(rowList(recordIndex))
                itemCount = CountOutputItems(sourceData, rowIndex, lastColumn)
                If itemCount > blockHeight Then blockHeight = itemCount
                columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP)
                valueIndex = 0

                For columnIndex = 1 To lastColumn
                    If ShouldOutput(sourceData(rowIndex, columnIndex)) Then
                        valueIndex = valueIndex + 1
                        Set outputCell = wsOutput.Cells(CONFIG_OUTPUT_START_ROW + rowOffset + valueIndex - 2, columnOffset)
                        outputCell.Interior.Pattern = headerFillPattern(columnIndex)
                        If headerFillPattern(columnIndex) <> xlNone Then outputCell.Interior.Color = headerFillColor(columnIndex)
                        outputCell.Font.Color = headerFontColor(columnIndex)
                    End If
                Next columnIndex
            End If
        Next blockPosition

        If CONFIG_BLOCK_ROW_GAP > 0 And blockRow < blockRows - 1 Then
            wsOutput.Rows(CONFIG_OUTPUT_START_ROW + rowOffset + blockHeight - 1).Resize(CONFIG_BLOCK_ROW_GAP).RowHeight = CONFIG_BLOCK_ROW_GAP_HEIGHT
        End If
        rowOffset = rowOffset + blockHeight + CONFIG_BLOCK_ROW_GAP
    Next blockRow

    If CONFIG_AUTOFIT_COLUMNS Then wsOutput.Columns.AutoFit
    With wsOutput.Cells(CONFIG_OUTPUT_START_ROW, 1).Resize(maxOutputRows, outputColumns)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    If CONFIG_BLOCK_COLUMN_GAP > 0 Then
        For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW - 1
            For gapIndex = 1 To CONFIG_BLOCK_COLUMN_GAP
                columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP) + 2 + gapIndex - 1
                wsOutput.Columns(columnOffset).ColumnWidth = CONFIG_BLOCK_COLUMN_GAP_WIDTH
            Next gapIndex
        Next blockPosition
    End If
    wsOutput.PageSetup.PrintArea = wsOutput.Cells(CONFIG_OUTPUT_START_ROW, 1).Resize(maxOutputRows, outputColumns).Address
End Sub

Private Sub ApplyPrintSetup(ByVal ws As Worksheet)
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
    End With
End Sub

Private Function CountOutputItems(ByRef sourceData As Variant, ByVal rowIndex As Long, ByVal lastColumn As Long) As Long
    Dim columnIndex As Long
    For columnIndex = 1 To lastColumn
        If ShouldOutput(sourceData(rowIndex, columnIndex)) Then CountOutputItems = CountOutputItems + 1
    Next columnIndex
End Function

Private Function ShouldOutput(ByVal value As Variant) As Boolean
    If IsError(value) Then
        ShouldOutput = True
    ElseIf CONFIG_SKIP_EMPTY_VALUES Then
        ShouldOutput = (Len(Trim$(CStr(value))) > 0)
    Else
        ShouldOutput = True
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    CellText = Trim$(CStr(value))
End Function

Private Function SafeSheetName(ByVal value As String) As String
    Dim invalidChars As Variant, item As Variant
    value = Trim$(value)
    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each item In invalidChars
        value = Replace(value, CStr(item), "-")
    Next item
    If Len(value) > 31 Then value = Left$(value, 31)
    SafeSheetName = value
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String, ByVal ignoredSheet As Worksheet) As String
    Dim candidate As String, suffix As String, index As Long
    Dim existing As Worksheet

    If Len(baseName) = 0 Then baseName = "Output"
    candidate = Left$(baseName, 31)
    index = 1

    Do
        Set existing = Nothing
        On Error Resume Next
        Set existing = wb.Worksheets(candidate)
        On Error GoTo 0

        If existing Is Nothing Then
            UniqueSheetName = candidate
            Exit Function
        End If
        If Not ignoredSheet Is Nothing Then
            If existing Is ignoredSheet Then
                UniqueSheetName = candidate
                Exit Function
            End If
        End If

        index = index + 1
        suffix = " (" & CStr(index) & ")"
        candidate = Left$(baseName, 31 - Len(suffix)) & suffix
    Loop
End Function
