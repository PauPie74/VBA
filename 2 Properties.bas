Sub Assign_value()

    'Cell A8 = 48
    Range("A8").Value = 58

End Sub


Sub Assign_value_another_workbook()

    'Cell A8 of sheet 2 of workbook 2 = Text example
    Workbooks("Book2.xlsx").Sheets("Sheet2").Range("A8").Value = "Text example"

End Sub


Sub Assign_value_range()

    '.Value part is not necessarily needed
    Range("A1:A8") = 48
    Range("B1:B8") = "text"
    Range("C1:C8") = 0 + 2

End Sub


Sub text_formating()

    'Change the text size of cells A1 to A8
    Range("A1:A8").Font.Size = 18
    
    Range("B1:B8").Font.Bold = True
    
    Range("A1:A8").Font.Italic = True
    
    Range("C1:C8").Font.Underline = True
    
    Range("B1:B8").Font.Name = "Arial"
    
End Sub


Sub Add_border()

    'Add a border to selected cells
    Selection.Borders.Value = 1
    
    'Value = 0: no border

End Sub


Sub hide_sheet()

    'Hide a sheet
    Sheets("Sheet3").Visible = 2

    'Visible = -1: display the sheet

End Sub


Sub take_value()

    'A7 = A1
    Range("A7") = Range("A1")

    'Or:
    'Range("A7").Value = Range("A1").Value
    
    'Copy text size
    
    Range("A6").Font.Size = Range("A2").Font.Size

End Sub


Sub counter()

    'Click counter in A1
    Range("A1") = Range("A1") + 1

End Sub

Sub properties_with()

    'Start of the instruction with: With to avoid repetition
    With Sheets("Sheet2").Range("A8")
        .Borders.Weight = 3
        With .Font
            .Bold = True
            .Size = 18
            .Italic = True
            .Name = "Arial"
        End With
    'End of the instruction with: End With
    End With
    
    'it's the same as
        'Sheets("Sheet2").Range("A8").Borders.Weight = 3
        'Sheets("Sheet2").Range("A8").Font.Bold = True
        'Sheets("Sheet2").Range("A8").Font.Size = 18
        'Sheets("Sheet2").Range("A8").Font.Italic = True
        'Sheets("Sheet2").Range("A8").Font.Name = "Arial"
End Sub