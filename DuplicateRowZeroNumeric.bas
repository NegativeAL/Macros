Sub DuplicateRowZeroNumeric()
    Dim ws As Worksheet
    Dim buttonRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim cell As Range
    Dim debugMsg As String
    Dim buttonName As String
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Determine which button was clicked
    buttonName = Application.Caller
    
    ' Find the button's row
    On Error Resume Next
    buttonRow = ws.Shapes(buttonName).TopLeftCell.Row
    On Error GoTo 0
    
    ' If the button row cannot be determined, exit the macro
    If buttonRow = 0 Then
        MsgBox "Unable to determine the button's row. Please ensure the button is properly placed.", vbExclamation
        Exit Sub
    End If
    
    ' Source row is always the row above the button
    sourceRow = buttonRow - 1
    
    ' Target row is always the button row
    targetRow = buttonRow
    
    ' Find the last column with data dynamically
    lastCol = ws.Cells(sourceRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Insert a new row
    ws.Rows(targetRow).Insert Shift:=xlDown
    
    ' After inserting the row, the button and all rows below it shift down by one row
    ' Update the buttonRow to reflect its new position
    buttonRow = buttonRow + 1
    
    ' Copy cells from the source row to the new row
    For i = 1 To lastCol
        ' Check if the cell in the source row exists
        If Not IsEmpty(ws.Cells(sourceRow, i)) Then
            ' Copy the cell to the new row
            ws.Cells(targetRow, i).Value = ws.Cells(sourceRow, i).Value
        End If
    Next i
    
    ' Zero out only numeric cells, preserving blank and text cells
    For i = 1 To lastCol
        Set cell = ws.Cells(targetRow, i)
        
        ' Check if the cell has a numeric value (ignore blank or text cells)
        If IsNumeric(cell.Value) And cell.Value <> "" Then
            cell.Value = 0
        End If
    Next i
    
    ' Extend sum formulas in the row below the button (buttonRow)
    debugMsg = "Debug Information:" & vbNewLine
    debugMsg = debugMsg & "Button Name: " & buttonName & vbNewLine
    debugMsg = debugMsg & "Button Row: " & buttonRow & vbNewLine
    debugMsg = debugMsg & "Last Column: " & lastCol & vbNewLine
    
    For i = 1 To lastCol
        Dim sumFormulaCell As Range
        Set sumFormulaCell = ws.Cells(buttonRow, i) ' Look in the row below the button
        
        ' Log the cell value and formula for debugging
        debugMsg = debugMsg & "Column " & i & " Cell Value: " & sumFormulaCell.Value & ", Formula: " & sumFormulaCell.Formula & vbNewLine
        
        ' Check if this cell is a SUM formula
        If Left(sumFormulaCell.Formula, 1) = "=" And _
           InStr(1, UCase(sumFormulaCell.Formula), "SUM(") > 0 Then
            
            ' Extract the current range
            Dim formulaRange As String
            formulaRange = Mid(sumFormulaCell.Formula, InStr(1, UCase(sumFormulaCell.Formula), "SUM(") + 4)
            formulaRange = Left(formulaRange, InStr(formulaRange, ")") - 1)
            
            ' Split the range
            Dim startCell As String
            Dim endCell As String
            startCell = Split(formulaRange, ":")(0)
            endCell = Split(formulaRange, ":")(1)
            
            ' Add range details to debug message
            debugMsg = debugMsg & "Start Cell: " & startCell & ", End Cell: " & endCell & vbNewLine
            
            ' Determine columns
            Dim startCol As Long
            Dim endCol As Long
            startCol = Range(startCell).Column
            endCol = Range(endCell).Column
            
            ' Determine new end row (the row just inserted)
            Dim newEndRow As Long
            newEndRow = buttonRow - 1 ' The new row is above the button row
            
            ' Construct new formula
            Dim newFormula As String
            newFormula = "=SUM(" & startCell & ":" & ws.Cells(newEndRow, endCol).Address(False, False) & ")"
            
            ' Add new formula to debug message
            debugMsg = debugMsg & "New Formula: " & newFormula & vbNewLine
            
            ' Attempt to update the formula
            On Error Resume Next
            sumFormulaCell.Formula = newFormula
            
            ' Check for errors
            If Err.Number <> 0 Then
                debugMsg = debugMsg & "Error updating formula: " & Err.Description & vbNewLine
                Err.Clear
            End If
            On Error GoTo 0
            
            ' Add final formula to debug message
            debugMsg = debugMsg & "Final Formula: " & sumFormulaCell.Formula & vbNewLine & vbNewLine
        End If
    Next i
    
    ' Display debug information
    ' MsgBox debugMsg
    
    ' Optional: Select the first cell of the new row
    ws.Cells(targetRow, 1).Select
End Sub
