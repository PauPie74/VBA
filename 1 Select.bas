Sub columnManipulations()

'Comment

    Range("B1:B18").Select
    Selection.ClearContents
    Range("C1:C18").Select
    Selection.ClearContents
    Range("D1:D18").Select
    Selection.Cut Destination:=Range("B1:B18")
    Range("B1:B18").Select
    
End Sub


Sub Select0()

    Sheets("2").Activate

    Range("A8").Select
    
    'Selection of the cell from row 8 and column 1
    Cells(8, 3).Select

End Sub


Sub Random_selection()

    'random_number_between_1_and_10

    Cells(Int(Rnd * 10) + 1, 1).Select

End Sub


Sub select_columns()

    'Selection of columns B to G
    Columns("B:G").Select

End Sub