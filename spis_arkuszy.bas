Sub spis_arkuszy()
  Const Spis = "Spis arkuszy"
  
  Dim sh As Worksheet
  Dim w As Integer
  
  'przejscie do arkusza ze spisem i ewentualne jego utworzenie
  On Error Resume Next
  Worksheets(Spis).Select
  If Err > 0 Then
    Worksheets.Add Before:=Sheets(1)
    ActiveSheet.Name = Spis
  End If
  On Error GoTo 0
  
  'wpisanie listy arkuszy
  Cells.Clear
  w = 0
  For Each sh In Worksheets
    w = w + 1
    Cells(w, 1) = sh.Name
  Next sh
End Sub