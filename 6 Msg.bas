Private Sub dialogBox(name As String, Optional firstName As String = "", Optional age As Integer = 0)

    'If age is missing
    If age = 0 Then
        
        If firstName = "" Then 'If the first name is missing, display only the name
            MsgBox name
        Else 'Otherwise, display the name and first name
            MsgBox name & " " & firstName
        End If

    'If age is provided
    Else

        If firstName = "" Then 'If the first name is missing, display the name and age
            MsgBox name & ", " & age & " years old"
        Else 'Otherwise, display the name, first name, and age
            MsgBox name & " " & firstName & ", " & age & " years old"
        End If
    
    End If
       
End Sub

Sub clearB2()

    If MsgBox("Are you sure you want to delete the contents of B2?", vbYesNo, "Confirmation") = vbYes Then
        Range("B2").ClearContents
        MsgBox "The content of B2 has been cleared!"
    End If

End Sub

Private Sub Workbook_Open()

    MsgBox "Welcome message"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    'If the user responds No, the Cancel variable will be set to True (which cancels the closure)
    If MsgBox("Are you sure you want to close this workbook?", 36, "Confirmation") = vbNo Then
        Cancel = True
    End If

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Static previousSelection As String

    'Remove background color from the previous selection
    If previousSelection <> "" Then
        Range(previousSelection).Interior.ColorIndex = xlColorIndexNone
    End If

    'Color the current selection
    Target.Interior.Color = RGB(181, 244, 0)

    'Save the address of the current selection
    previousSelection = Target.Address

End Sub
