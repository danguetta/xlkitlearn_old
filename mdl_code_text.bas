Attribute VB_Name = "mdl_code_text"
Option Explicit

Sub load_code()
    Dim original_calc_mode
    original_calc_mode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' Read the code from the xlkitlearn.py file
    If FileExists(ThisWorkbook.Path & "/xlkitlearn.py") Then
        Open ThisWorkbook.Path & "/xlkitlearn.py" For Input As #1
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

Function output_files()
    Dim Export As Boolean
    Dim Source As Excel.Workbook
    Dim SourceWorkbook As String
    Dim ExportPath As String
    Dim FileName As String
    ' Must include Microsoft Visual Basic for Applications Extensibility Reference
    Dim Component As VBIDE.VBComponent

    ' This workbook must be open in Excel.
    SourceWorkbook = ActiveWorkbook.Name
    Set Source = Application.Workbooks(SourceWorkbook)
    
    ' Must update security settings to trust access to the VBA project object model
    If Source.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Function
    End If
    
    #If Mac Then
        ExportPath = ThisWorkbook.Path & "/"
    #Else
        ExportPath = ThisWorkbook.Path & "\"
    #End If
    
    For Each Component In Source.VBProject.VBComponents

        Export = True
        FileName = Component.Name

        ' Create filenames and extensions
        Select Case Component.Type
            Case vbext_ct_ClassModule
                FileName = FileName & ".cls"
            Case vbext_ct_MSForm
                FileName = FileName & ".frm"
            Case vbext_ct_StdModule
                FileName = FileName & ".bas"
            Case vbext_ct_Document
                If FileName = "Sheet1" Or FileName = "Sheet2" Or FileName = "ThisWorkbook" Then
                    FileName = FileName & ".sht"
                Else
                    ' This is a worksheet or workbook object with no code. Don't export
                    Export = False
                End If
        End Select
        
        If Export Then
            ' Export the component
            If FileExists(ExportPath & FileName) Then
                Kill ExportPath & FileName
            End If
            
            Component.Export ExportPath & FileName
        End If
   
    Next Component
    
    ' Delete all files with frx extension
    Kill ThisWorkbook.Path & "/*.frx*"
    
End Function
