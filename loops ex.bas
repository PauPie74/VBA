Sub loopsExercise()

    Dim column As Integer
    Dim row As Integer
    Dim number As Integer
    
    For row = 1 To 10:
        For column = 1 To 10
            number = column * row
            Cells(row, column) = number
            If (row + column) Mod 2 = 0 Then
                Cells(row, column).Interior.Color = RGB(220, 220, 220)
            End If
        Next
    Next
    
End Sub
