Attribute VB_Name = "AutoCable"
Option Explicit

' CONFIG
Private Const CONFIG_HEADER_ROW As Long = 1
Private Const CONFIG_START_ROW As Long = 3
Private Const CONFIG_GROUP_COLUMN As Long = 1
Private Const CONFIG_BLOCKS_PER_ROW As Long = 4
Private Const CONFIG_OUTPUT_START_ROW As Long = 1
Private Const CONFIG_OUTPUT_START_COLUMN As Long = 2
Private Const CONFIG_BLOCK_COLUMN_GAP As Long = 1
Private Const CONFIG_BLOCK_ROW_GAP As Long = 1
Private Const CONFIG_BLOCK_ROW_GAP_HEIGHT As Double = 15
Private Const CONFIG_BLOCK_COLUMN_GAP_WIDTH As Double = 2
Private Const CONFIG_SKIP_EMPTY_VALUES As Boolean = True
Private Const CONFIG_DRAW_BORDERS As Boolean = True
Private Const CONFIG_AUTOFIT_COLUMNS As Boolean = True

Public Sub ProcessLocationSheets()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim sourceData As Variant
    Dim groups As Object
    Dim key As Variant
    Dim lastRow As Long, lastColumn As Long
    Dim rowIndex As Long, columnIndex As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim overwriteMode As Long
    Dim targetName As String
    Dim userChoice As Variant
    Dim savedError As String

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Select a worksheet and try again.", vbExclamation, "AutoCable"
        Exit Sub
    End If

    Set wsSource = ActiveSheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, CONFIG_GROUP_COLUMN).End(xlUp).Row
    lastColumn = wsSource.Cells(CONFIG_HEADER_ROW, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRow < CONFIG_START_ROW Or lastColumn < 1 Then
        MsgBox "No data was found in the configured range.", vbInformation, "AutoCable"
        Exit Sub
    End If

    sourceData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastColumn)).Value2
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    For rowIndex = CONFIG_START_ROW To UBound(sourceData, 1)
        Dim groupName As String
        groupName = CellText(sourceData(rowIndex, CONFIG_GROUP_COLUMN))
        If Len(groupName) > 0 Then
            If Not groups.Exists(groupName) Then groups.Add groupName, New Collection
            groups(groupName).Add rowIndex
        End If
    Next rowIndex

    If groups.Count = 0 Then
        MsgBox "No non-empty group values were found in the configured column.", vbInformation, "AutoCable"
        Exit Sub
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
    overwriteMode = 0

    For Each key In groups.Keys
        targetName = SafeSheetName(CStr(key))
        If Len(targetName) = 0 Then GoTo NextGroup

        Set wsOutput = Nothing
        On Error Resume Next
        Set wsOutput = wsSource.Parent.Worksheets(targetName)
        On Error GoTo ErrorHandler

        If Not wsOutput Is Nothing Then
            If wsOutput Is wsSource Then
                MsgBox "The output name matches the source sheet: " & targetName & vbCrLf & _
                       "This group was skipped to protect the source data.", vbExclamation, "AutoCable"
                GoTo NextGroup
            End If

            If overwriteMode = 0 Then
                userChoice = Application.InputBox( _
                    Prompt:="A sheet named '" & targetName & "' already exists." & vbCrLf & vbCrLf & _
                            "1 = Skip this group" & vbCrLf & _
                            "2 = Override this sheet" & vbCrLf & _
                            "3 = Skip all remaining conflicts" & vbCrLf & _
                            "4 = Override all remaining conflicts", _
                    Title:="AutoCable - Existing Sheet", Type:=1)

                If VarType(userChoice) = vbBoolean And userChoice = False Then GoTo Cleanup
                If Not IsNumeric(userChoice) Or userChoice < 1 Or userChoice > 4 Or userChoice <> Fix(userChoice) Then
                    MsgBox "Enter 1, 2, 3, or 4. Processing was cancelled.", vbExclamation, "AutoCable"
                    GoTo Cleanup
                End If

                Select Case CLng(userChoice)
                    Case 1: overwriteMode = 1
                    Case 2: overwriteMode = 2
                    Case 3: overwriteMode = 1
                    Case 4: overwriteMode = 2
                End Select

                If CLng(userChoice) = 1 Or CLng(userChoice) = 3 Then GoTo NextGroup
            ElseIf overwriteMode = 1 Then
                GoTo NextGroup
            End If

            wsOutput.Delete
            Set wsOutput = Nothing
        End If

        Set wsOutput = wsSource.Parent.Worksheets.Add(After:=wsSource.Parent.Worksheets(wsSource.Parent.Worksheets.Count))
        wsOutput.Name = targetName
        WriteGroupOutput wsOutput, wsSource, sourceData, groups(key), lastColumn

NextGroup:
    Next key

    MsgBox "Processing completed.", vbInformation, "AutoCable"
    GoTo Cleanup

ErrorHandler:
    savedError = Err.Description
    MsgBox "Processing failed: " & savedError, vbCritical, "AutoCable"

Cleanup:
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
                    If CountOutputItems(sourceData, rowIndex, lastColumn) > 0 Then
                        columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP)
                        With wsOutput.Range(wsOutput.Cells(CONFIG_OUTPUT_START_ROW + rowOffset - 1, columnOffset), _
                                            wsOutput.Cells(CONFIG_OUTPUT_START_ROW + rowOffset + CountOutputItems(sourceData, rowIndex, lastColumn) - 2, columnOffset + 1))
                            .Borders.LineStyle = xlContinuous
                        End With
                    End If
                End If
            Next blockPosition
        End If
        rowOffset = rowOffset + blockHeight + CONFIG_BLOCK_ROW_GAP
    Next blockRow

    wsOutput.Cells(CONFIG_OUTPUT_START_ROW, 1).Resize(maxOutputRows, outputColumns).Value = outputData

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
    If CONFIG_BLOCK_COLUMN_GAP > 0 Then
        For blockPosition = 1 To CONFIG_BLOCKS_PER_ROW - 1
            For gapIndex = 1 To CONFIG_BLOCK_COLUMN_GAP
                columnOffset = CONFIG_OUTPUT_START_COLUMN + (blockPosition - 1) * (2 + CONFIG_BLOCK_COLUMN_GAP) + 2 + gapIndex - 1
                wsOutput.Columns(columnOffset).ColumnWidth = CONFIG_BLOCK_COLUMN_GAP_WIDTH
            Next gapIndex
        Next blockPosition
    End If
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
