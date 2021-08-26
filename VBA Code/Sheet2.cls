Private Sub Worksheet_Activate()
    If DEBUG_VERSION = False Then
        ThisWorkbook.Sheets("code_text").Visible = 2
    End If
End Sub