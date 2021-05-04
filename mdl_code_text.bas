Option Explicit

Sub load_code()
    Dim original_calc_mode
    original_calc_mode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' Read the code from the xlkitlearn.py file
    If FileExists(ThisWorkbook.Path & "\xlkitlearn.py") Then
        Open ThisWorkbook.Path & "\xlkitlearn.py" For Input As #1
    Else
        MsgBox "Could not find code file xlkitlearn.py.", vbCritical
        Exit Sub
    End If
    
    Dim code_text As String
    code_text = Input(LOF(1), 1)
    
    Close #1
    
    ' Split on newlines
    Dim split_text As Variant
    split_text = Split(code_text, vbLf)
    
    ' Clear the worksheet
    ThisWorkbook.Sheets("code_text").Cells.Delete
    
    ' Output the number of lines to the first row
    ThisWorkbook.Sheets("code_text").Range("A1") = UBound(split_text) + 1
        
    ' Output the rest of the code
    Dim i As Integer
    For i = 0 To UBound(split_text)
        ThisWorkbook.Sheets("code_text").Range("A" & (i + 2)) = split_text(i)
    Next i
    
    update_names
    
    Application.Calculation = original_calc_mode
End Sub
