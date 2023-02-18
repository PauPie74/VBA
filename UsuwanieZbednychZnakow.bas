Sub UsuwanieZnakow()

    'deklaracja zmiennych
    Dim i As Integer, komorka As Object
    Dim tekst As String
    Dim ListaZnakow As String
    Dim znaki
    
    'Pobranie listy znaków od użytkownika
    ListaZnakow = InputBox("Wpisz znaki do usunięcia (znaki rozdzielaj spacją)", "Lista znaków")
    
    
    'deklarowanie tablic z literami
    znaki = Split(ListaZnakow, " ")
   
    'pętla przechodząca po komórkach z zaznaczonego zakresu
    For Each komorka In Selection
        'pobranie wartości komórki do zmiennej
        tekst = komorka.Value
        'pętla szukająca znakow do usuniecia
        For i = 0 To UBound(znaki)
            'zamiana znaków
            tekst = WorksheetFunction.Substitute(tekst, znaki(i), "")
        Next i
        'zwrócenie przetworzonej zawartosci komorki
        komorka.Value = tekst
    Next komorka

End Sub