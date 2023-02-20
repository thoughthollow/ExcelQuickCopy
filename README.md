# ExcelQuickCopy

VBA code.
If the ActiveCell is in a certain range, then colour the cell AND copy the cell AND minimise the window.


Note to self:

*Upload later with the below. It checks whether the cell has a value and wont act on empty cells.*
*(also maybe look into whether there's an easy way to select the effected range rather than manually altering the script).*

 Private Sub worksheet_selectionchange(ByVal target As Range)

  If Not Intersect(ActiveCell, Range("B2:B9999")) Is Nothing Then

   If Not IsEmpty(ActiveCell.Value) Then

     ActiveCell.Interior.ColorIndex = 20
     ActiveCell.Copy
     ActiveWindow.WindowState = xlMinimized

   End If

  End If

 End Sub
