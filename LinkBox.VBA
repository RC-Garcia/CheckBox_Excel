Sub LinkCheckBoxes()
Dim chk As CheckBox
Dim lCol As Long
lCol = 0 'number of columns to the right of checkbox
lRow = 1
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(lRow, lCol).Address
   End With
Next chk
End Sub
