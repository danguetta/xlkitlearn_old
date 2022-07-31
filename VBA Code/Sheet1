Option Explicit

Dim change_running As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    Target.Hyperlinks.Delete
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    If Not change_running Then
        change_running = True
        
        If Target.Address = "$D$8" Then
            ActiveSheet.Range("F17").Select
        ElseIf Target.Address = "$D$10" Then
            ActiveSheet.Range("D14").Select
        ElseIf Target.Address = "$D$13" Then
            ActiveSheet.Range("D9").Select
        ElseIf Target.Address = "$D$15" Then
            ActiveSheet.Range("F17").Select
        ElseIf Target.Address = "$F$16" Then
            ActiveSheet.Range("D14").Select
        ElseIf Target.Address = "$F$18" Then
            ActiveSheet.Range("D9").Select
        ElseIf Target.row <= 9 Then
            ActiveSheet.Range("D9").Select
        ElseIf Target.row <= 14 Then
            ActiveSheet.Range("D14").Select
        Else
            ActiveSheet.Range("F17").Select
        End If
        
        change_running = False
    End If
End Sub