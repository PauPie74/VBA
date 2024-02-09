Type User
    name As String
    firstName As String
End Type

Sub variables()

    'Declaration of the variable
    Dim myVariable As Integer
    
    'Attribution of a value to the variable
    myVariable = 12
    
    'Display of the value of myVariable in a MsgBox
    MsgBox myVariable
    
End Sub

Sub varSheet()

    Dim varSheet As Worksheet
    Set varSheet = Sheets("Sheet2")
    varSheet.Activate
    
    Dim example As Integer
    Dim example%
    'symbold are used to shorten the delcaration, % - integar, @ - currency, $ - string etc...

End Sub

Sub cells_as_vars()

    Dim name As String, firstName As String, age As Integer

    name = Cells(2, 1)
    firstName = Cells(2, 2)
    age = Cells(2, 3)

    MsgBox name & " " & firstName & ", " & age & " years old"

End Sub

Sub cells_as_vars2()

    MsgBox Cells(Range("F5") + 1, 1) & " " & Cells(Range("F5") + 1, 2) & ", " & Cells(Range("F5") + 1, 3) & " years old"

End Sub

Sub array_Dec()

    'Automattically it's 1-dim array: 0, 1, 2... 4
    Dim array1(4) As String
    
    'declaring values
    array1(0) = "Value of cell 0"
    array1(1) = "Value of cell 1"
    array1(2) = "Value of cell 2"
    array1(3) = "Value of cell 3"
    array1(4) = "Value of cell 4"
    
End Sub


'Declaring own type of variable ^ Type is at the beginning

Sub create_user()

    'Declaration
    Dim user1 As User
    
    'Assigning values to user1
    user1.name = "Smith"
    user1.firstName = "John"
    
    'Example of use
    MsgBox user1.name & " " & user1.firstName

End Sub
