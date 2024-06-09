' in Excel VBA, offset a range robustly, in a way that ignores merged cells

Sub testCustomOffset()

    MsgBox customOffset(Range("A1:B3"), 0, 1).Address
    
End Sub

Function customOffset(rng As Range, rowOffset As Integer, colOffset As Integer) As Range

    Dim firstCell As Range
    Set firstCell = rng.Cells(1, 1)
    
    Dim lastCell As Range
    Set lastCell = rng.Cells(rng.Rows.Count, rng.Columns.Count)
    
    Set customOffset = Range(firstCell.Cells(rowOffset + 1, colOffset + 1), lastCell.Cells(rowOffset + 1, colOffset + 1))

End Function
