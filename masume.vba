Option Explicit

Sub MoveSelection(offsetRow As Integer, offsetColumn As Integer)
    Dim selectedRange As Range
    Dim tempRange As Range
    Dim tempRangeArr As Variant
    Dim movedRange As Range
    
    On Error Resume Next
    Set selectedRange = Selection
    On Error GoTo 0
    
    If Not selectedRange Is Nothing Then
        ' Create a temporary array to store the contents of the moving cells
        tempRangeArr = GetRange(selectedRange)
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' MOVE UP OR DOWN '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If offsetColumn = 0 Then
            ' Get the coordinates of the destination range
            Dim destRow As Long
            Dim destColumn As Long
            destRow = selectedRange.Cells(1, 1).row + offsetRow
            destColumn = selectedRange.Cells(1, 1).column + offsetColumn
            
            ' Check if the destination cell is within the worksheet bounds
            If IsValidCell(destRow, destColumn) Then
                ' Define the destination range
                Dim destinationRange As Range
                Set destinationRange = selectedRange.Offset(offsetRow, 0)
                
                ' Clear the selectedRange
                selectedRange.Clear
                
                ' Check if there are any data within the destination range
                If Not IsEmpty(destinationRange.Value) Then
                    ' If there is content, move the content values down or above the destination range
                    If offsetRow < 0 Then
                        destinationRange.Offset(selectedRange.Rows.Count, 0).Resize(1, selectedRange.Columns.Count).Value = destinationRange.Value
                    Else
                        destinationRange.Rows(destinationRange.Rows.Count).Offset(-selectedRange.Rows.Count, 0) = destinationRange.Rows(destinationRange.Rows.Count).Value
                    End If
                End If
                
                ' Paste the temporary array values into the destination range
                Dim i, j
                For i = LBound(tempRangeArr, 1) To UBound(tempRangeArr, 1)
                    For j = LBound(tempRangeArr, 2) To UBound(tempRangeArr, 2)
                        Cells(destRow, destColumn) = tempRangeArr(i, j)
                        destColumn = destColumn + 1
                    Next j
                    destColumn = destinationRange.Cells(1, 1).column
                    destRow = destRow + 1
                Next i
                
                Set movedRange = destinationRange
            End If
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' MOVE LEFT OR RIGHT '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If offsetRow = 0 Then
            ' Get the coordinates of the destination range
            destRow = selectedRange.Cells(1, 1).row + offsetRow
            destColumn = selectedRange.Cells(1, 1).column + offsetColumn
            
            ' Check if the destination cell is within the worksheet bounds
            If IsValidCell(destRow, destColumn) Then
                ' Define the destination range
                Set destinationRange = selectedRange.Offset(0, offsetColumn)
                
                ' Check if selectedRange cell count is one
                If selectedRange.Rows.Count + selectedRange.Columns.Count = 2 Then
                    ' GOING LEFT
                    If offsetColumn < 0 Then
                        ' Shift the cell to the left by one cell
                        destinationRange.Columns(1).Insert xlShiftDown
                        ' Delete the cell to the right where the value will leave
                        destinationRange.Offset(-1, 1).Resize(1, 1).Delete Shift:=xlUp
                        ' Reset the destination range for the copy
                        Set destinationRange = destinationRange.Offset(-1, 0)
                    ' GOING RIGHT
                    Else
                        ' Shift the cell to the right by one cell
                        destinationRange.Columns(destinationRange.Columns.Count).Insert xlShiftDown
                        ' Delete the cell to the left where the value will leave
                        destinationRange.Offset(-1, -1).Resize(1, 1).Delete Shift:=xlUp
                        ' Reset the destination range for the copy
                        Set destinationRange = destinationRange.Offset(-1, 0)
                    End If
                    
                ' selectedRange is multi cell
                Else
                    selectedRange.Clear
                    ' GOING LEFT
                    If offsetColumn < 0 Then
                        ' Shift cells to the left, down by the destinationRange's rows
                        destinationRange.Columns(1).Insert xlShiftDown
                        ' Reset destination range
                        Set destinationRange = selectedRange.Offset(0, offsetColumn)
                        ' Shift cells to the right, up by the number of destinationRange's rows
                        destinationRange.Offset(0, destinationRange.Columns.Count).Resize(destinationRange.Rows.Count, 1).Delete Shift:=xlUp
                    ' GOING RIGHT
                    Else
                        ' Shift cells to the right, down by the destinationRange's rows
                        destinationRange.Columns(destinationRange.Columns.Count).Insert xlShiftDown
                        ' Reset destination range
                        Set destinationRange = selectedRange.Offset(0, offsetColumn)
                        ' Shift cells to the left, up by the number of destinationRange's rows
                        destinationRange.Offset(0, -1).Resize(destinationRange.Rows.Count, 1).Delete Shift:=xlUp
                    End If
                End If
                
                ' Paste the temporary array values into the destination range
                For i = LBound(tempRangeArr, 1) To UBound(tempRangeArr, 1)
                    For j = LBound(tempRangeArr, 2) To UBound(tempRangeArr, 2)
                        Cells(destRow, destColumn) = tempRangeArr(i, j)
                        destColumn = destColumn + 1
                    Next j
                    destColumn = destinationRange.Cells(1, 1).column
                    destRow = destRow + 1
                Next i
                
                Set movedRange = destinationRange
            End If
        End If

        ' Select the range of moved cells
        If Not movedRange Is Nothing Then
            movedRange.Select
        End If
    End If
End Sub

Function IsValidCell(row As Long, column As Long) As Boolean
    ' Check if the cell is within the worksheet bounds
    IsValidCell = row > 0 And row <= Rows.Count And column > 0 And column <= Columns.Count
End Function

Function IsValidRange(rng As Range) As Boolean
    On Error Resume Next
    ' Attempt to access a property of the range
    Dim checkValue As Variant
    checkValue = rng.Cells(1, 1).Value
    On Error GoTo 0
    
    ' Check if no error occurred (i.e., the range is valid)
    IsValidRange = (Err.Number = 0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Returns the values of a range ('rg') in a 2D one-based array.
' Remarks:      If ˙rg` refers to a multi-range, only its first area
'               is considered.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetRange( _
    ByVal rg As Range) _
As Variant
    Const ProcName As String = "GetRange"
    On Error GoTo ClearError
    
    If rg.Rows.Count + rg.Columns.Count = 2 Then ' one cell
        Dim Data As Variant: ReDim Data(1 To 1, 1 To 1): Data(1, 1) = rg.Value
        GetRange = Data
    Else ' multiple cells
        GetRange = rg.Value
    End If

ProcExit:
    Exit Function
ClearError:
    Debug.Print "'" & ProcName & "' Run-time error '" _
        & Err.Number & "':" & vbLf & "    " & Err.Description
    Resume ProcExit
End Function

Sub MoveSelectionUp()
    MoveSelection -1, 0
End Sub

Sub MoveSelectionDown()
    MoveSelection 1, 0
End Sub

Sub MoveSelectionLeft()
    MoveSelection 0, -1
End Sub

Sub MoveSelectionRight()
    MoveSelection 0, 1
End Sub

Sub AssignShortcuts()
    ' Assign shortcuts to the macros
    Application.OnKey "+^{UP}", "MoveSelectionUp"    ' Ctrl + Alt + Up Arrow
    Application.OnKey "+^{DOWN}", "MoveSelectionDown"  ' Ctrl + Alt + Down Arrow
    Application.OnKey "+^{LEFT}", "MoveSelectionLeft"  ' Ctrl + Alt + Left Arrow
    Application.OnKey "+^{RIGHT}", "MoveSelectionRight" ' Ctrl + Alt + Right Arrow
End Sub
