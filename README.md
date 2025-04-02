# Module-01

Sub PasteToFilteredCells()
    Dim pasteRange As Range
    Dim copyData As Variant
    Dim i As Long, pasteRow As Range
    
    ' Get copied data from clipboard
    On Error GoTo ErrHandler
    copyData = Application.WorksheetFunction.Transpose(Application.Clipboard.GetText)

    ' Confirm selection is valid
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the filtered column you want to paste into.", vbExclamation
        Exit Sub
    End If

    Set pasteRange = Selection.SpecialCells(xlCellTypeVisible)

    If UBound(copyData) + 1 <> pasteRange.Cells.Count Then
        MsgBox "The number of items you're trying to paste doesn't match the number of visible cells.", vbExclamation
        Exit Sub
    End If

    ' Paste values to filtered cells
    i = 1
    For Each pasteRow In pasteRange.Cells
        pasteRow.Value = copyData(i)
        i = i + 1
    Next pasteRow

    MsgBox "Data pasted successfully to visible (filtered) cells.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error occurred. Make sure you copied data before running this macro.", vbCritical
End Sub
