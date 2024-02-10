Sub exercise()
    
    Dim lastRow As Integer, search As String, years As Integer, DByear As Integer, CustomerID As Integer, DBpay As String, currValue As Integer
    
    'Last row of the database
    lastRow = Sheets("DB").Cells(Rows.Count, 1).End(xlUp).Row

    'Value to search for (YES or NO)
    If Sheets("Grid").OptionButton_yes Then
        search = "YES"
    Else
        search = "NO"
    End If

    Dim dataArray(30)
    
    For arra = 0 To 30:
        dataArray(arra) = 0
    Next

    For i = 2 To 17
        years = Sheets("Grid").Cells(i, 1).Value
        
        For j = 2 To lastRow
            DByear = Year(Sheets("DB").Cells(j, 1).Value)
            CustomerID = Sheets("DB").Cells(j, 2).Value
            DBpay = Sheets("DB").Cells(j, 3).Value
            
            If DByear = years Then
                If DBpay = search Then
                    currValue = dataArray(CustomerID)
                    dataArray(CustomerID) = currValue + 1
                End If
            End If
        Next
        
        For k = 2 To 31
            Sheets("Grid").Cells(i, k) = dataArray(k - 1)
        Next
        
        For arr = 0 To 30:
            dataArray(arr) = 0
        Next
    Next

End Sub
