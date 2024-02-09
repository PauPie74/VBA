Sub color_font()

    'Text color in A1: blue (color 23)
    Range("A1").Font.ColorIndex = 23

End Sub

Sub color_font_ver2()

    'Text color in A1: RGB(0, 102, 204)
    Range("A2").Font.Color = RGB(0, 102, 204)

End Sub


Sub color_border()

    'Border thickness
    ActiveCell.Borders.Weight = 4

    'Border color: red
    ActiveCell.Borders.Color = RGB(255, 0, 0)

End Sub

Sub colors_inside()

    'Color the background of the selected cells
    Selection.Interior.Color = RGB(215, 238, 247)

End Sub

Sub colors_sheetTab()

    'Color the tab of the sheet "Sheet1"
    Sheets("Sheet2").Tab.Color = RGB(255, 0, 0)

End Sub
