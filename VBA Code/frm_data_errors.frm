Private Sub UserForm_Initialize()
    lst_errors.ColumnCount = 5
    lst_errors.ColumnWidths = "50,15,170,15,500"
    
    Dim i As Integer
    For i = 1 To header_errors.count()
        lst_errors.AddItem
        
        lst_errors.List(i - 1, 0) = Split(header_errors.Item(i), "`")(0)
        lst_errors.List(i - 1, 1) = ""
        lst_errors.List(i - 1, 2) = Split(header_errors.Item(i), "`")(1)
        lst_errors.List(i - 1, 3) = ""
        lst_errors.List(i - 1, 4) = Split(header_errors.Item(i), "`")(2)
    Next i
    
End Sub