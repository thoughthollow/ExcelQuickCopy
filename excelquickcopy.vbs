Private Sub worksheet_selectionchange(ByVal target As Range)

 If Not Intersect(ActiveCell, Range("A1:Z99")) Is Nothing Then
  
  ActiveCell.Interior.ColorIndex = 20
  ActiveCell.Copy
  ActiveWindow.WindowState = xlMinimized
 
 End If
 
End Sub
