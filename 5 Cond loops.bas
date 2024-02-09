Sub if_example()

    'If F5 is numeric
    If IsNumeric(Range("F5")) Then

        Dim name As String, firstName As String, age As Integer, lineNumber As Integer, nbRows As Integer
        
        lineNumber = Range("F5") + 1
        nbRows = WorksheetFunction.CountA(Range("A:A")) 'COUNTA function

        'If the number is in the correct range
        If lineNumber >= 2 And lineNumber <= nbRows Then
            name = Cells(lineNumber, 1)
            firstName = Cells(lineNumber, 2)
            age = Cells(lineNumber, 3)
            MsgBox name & " " & firstName & ", " & age & " years old"

        'If the number is out of range
        Else
            MsgBox "The entry """ & Range("F5") & """ is not a valid number!"
            Range("F5") = ""
        End If

    'If F5 is not numeric
    Else
        MsgBox "The entry """ & Range("F5") & """ is not valid!"
        Range("F5") = ""
    End If
    
End Sub

Sub comments()

    'Variables
    Dim grade As Single, comment As String
    grade = Range("A1")
    
    'Comment based on the grade
    Select Case grade '<= the value to test (here, the grade)
        Case Is = 6
            comment = "Excellent result!"
        Case Is >= 5
            comment = "Good result"
        Case Is >= 4
            comment = "Satisfactory result"
        Case Is >= 3
            comment = "Unsatisfactory result"
        Case Is >= 2
            comment = "Bad result"
        Case Is >= 1
            comment = "Terrible result"
        Case Else
            comment = "No result"
    End Select
    
    'Comment in B1
    Range("B1") = comment

End Sub

Sub is_empty_example()
    
    Dim myVariable
    
    If IsEmpty(myVariable) Then
        MsgBox "My variable hasn't been initialized!"
    Else
        MsgBox "My variable contains: " & myVariable
    End If

End Sub

'The * character can replace: no character, one character, or multiple characters:
' If myVariable Like "*12345*" Then '=> True

' The # character can replace a numeric character from 0 to 9:
' If myVariable Like "Example 12###" Then '=> True

'The ? character can replace any character:
'If myVariable Like "?xample?1234?" Then '=> True


Sub loop_while()

    Dim number As Integer

    number = 1 'Starting number

    Do While number <= 12 'While the variable "number" is <= 12, the loop is repeated
        Cells(number, 1) = number 'Numbering
        number = number + 1 'The number is increased by 1 at each loop
    Loop
    
    'Instead of repeating the loop while the condition is true, it's possible to exit the loop when the condition is true by replacing While with Until

End Sub


Sub loop_for()

    Dim i As Integer

    For i = 1 To 5
        MsgBox i 'Outputs the values: 1 / 2 / 3 / 4 / 5
    Next

End Sub

Sub loop_each()

    Dim sheet As Worksheet
    
    For Each sheet In Worksheets
        MsgBox sheet.name
    Next

End Sub

