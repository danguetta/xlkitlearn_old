Option Explicit

Dim start_time As Long

Dim sheet_names As New Collection
Dim sheet_settings As New Collection
Dim running_batch_process As Boolean

' =========================
' = Show settings buttons =
' =========================

Public Function validate_settings_string(which_addin As String, allow_blank As Boolean) As Boolean
    Dim settings_cell As String
    Dim blank_settings_string As String
    
    If which_addin = "predictive" Then
        settings_cell = CURRENT_SETTINGS
        blank_settings_string = BLANK_SETTINGS
    ElseIf which_addin = "text" Then
        settings_cell = CURRENT_TEXT_SETTINGS
        blank_settings_string = BLANK_TEXT_SETTINGS
    End If
    
    validate_settings_string = False
    If Trim(Range(settings_cell).Value) = "" Then
        If allow_blank Then
            Range(settings_cell).Value = blank_settings_string
            validate_settings_string = True
        Else
            MsgBox "It look like your settings string (cell " & Split(settings_cell, "!")(1) & ") is blank; Please enter some settings and try again.", vbCritical
        End If
    ElseIf UBound(dict_utils(settings_cell)) = -1 Then
        MsgBox "It look like you messed with your settings string (cell " & Split(settings_cell, "!")(1) & "); try clearing it and starting again.", vbCritical
    Else
        validate_settings_string = True
    End If
End Function

Sub btn_edit_settings()
    ' If code isn't running, show settings
    If run_id() = "" And validate_settings_string("predictive", True) Then
        update_conf True
        frm_pred.Show False
    End If
End Sub

Sub btn_edit_text_settings()
    ' If code isn't running, show settings
    If run_id() = "" And validate_settings_string("text", True) Then
        update_conf True
        frm_text.Show False
    End If
End Sub

' ====================================
' = Control enabled/disabled buttons =
' ====================================

Sub disable_buttons()
    ThisWorkbook.Sheets("Add-in").Buttons("btn_run").Font.Color = 8421504
    ThisWorkbook.Sheets("Add-in").Buttons("btn_run_text").Font.Color = 8421504
    
    ThisWorkbook.Sheets("Add-in").Buttons("btn_edit").Font.Color = 8421504
    ThisWorkbook.Sheets("Add-in").Buttons("btn_edit_text").Font.Color = 8421504
End Sub

Sub enable_buttons()
    ThisWorkbook.Sheets("Add-in").Buttons("btn_run").Font.Color = 0
    ThisWorkbook.Sheets("Add-in").Buttons("btn_run_text").Font.Color = 0
    
    ThisWorkbook.Sheets("Add-in").Buttons("btn_edit").Font.Color = 0
    ThisWorkbook.Sheets("Add-in").Buttons("btn_edit_text").Font.Color = 0
End Sub

' ======================
' = Run add-in buttons =
' ======================

Sub btn_run_addin()
    ' If the add-in is not running, run it
    If run_id() = "" And validate_settings_string("predictive", False) Then
        ' If any entries are missing from the dictionary, sync it
        Range(CURRENT_SETTINGS).Value = sync_dicts(Range(CURRENT_SETTINGS).Value, BLANK_SETTINGS)
        
        ' Run
        run_addin "predictive_addin", STATUS_CELL
    End If
End Sub

Sub btn_run_text_addin()
    ' If the add-in is not running, run it
    If run_id() = "" And validate_settings_string("text", False) Then
        ' If any entries are missing from the dictionary, sync it
        Range(CURRENT_TEXT_SETTINGS).Value = sync_dicts(Range(CURRENT_TEXT_SETTINGS).Value, BLANK_TEXT_SETTINGS)
        
        ' Run
        run_addin "text_addin", STATUS_CELL_TEXT
    End If
End Sub

Public Sub kill_addin()
    Dim code_text As String

    code_text = "import xlwings as xw; import types; mod = types.ModuleType('mod'); " & _
            "code_sheet = xw.Book.caller().sheets('code_text'); " & _
            "code_range = 'A2:A' + str(int(code_sheet.range('A1').value+1)); " & _
            "exec('\n'.join([str(i) if i is not None else '' for i in code_sheet.range(code_range).value]), mod.__dict__); " & _
            "mod.kill_addin()"
    
    Dim run_python_result As String
    run_python_result = RunPython(code_text)
        
    'Dim kill_code As String
   '
   ' kill_code = "import xlwings as xw\n" & _
   '             "import os\n" & _
   '             "import signal\n" & _
   '             "pid = xw.Book.caller().Sheets('code_text').Range('C1').value\n" & _
   '             "try:\n" & _
   '             "    os.kill(int(pid), signal.SIGTERM)\n" & _
   '             "except:\n" & _
   '             "    pass\n" & _
   '             "xw.Book.caller().macro('format_sheet')()"
   '
   ' kill_code = Replace(kill_code, "'", "\'")
   '
   ' RunPython kill_code
End Sub

' ==============
' = Run add-in =
' ==============

Sub run_addin(f_name As String, this_status_cell As String)
    ' Load code if it's here and if we're not in production
    If PROD_VERSION = False Then
        On Error Resume Next
        load_code
        resolve_workbook_path
        On Error GoTo 0
    End If

    ' Ensure an email was provided
    Dim email_provided As String
    email_provided = ActiveWorkbook.Sheets("Add-in").Range("F17").Value
    If Trim(email_provided) = "" Then
        MsgBox "Please enter an email address on the addin tab to validate the add-in before it runs.", vbExclamation, "Validation error"
        format_sheet
        Exit Sub
    End If
    
    ' If data is in a separate file, check that it exists
    Dim file_path As String: file_path = ""
    If f_name = "predictive_addin" Then
        If Left(frm_pred.lbl_training_data.tag, 5) = "File:" Then
            file_path = Right(frm_pred.lbl_training_data.tag, 7)
        End If
    ElseIf f_name = "text_addin" Then
        file_path = frm_text.txt_source_data.Text
    End If
    
    If file_path <> "" Then
        #If Mac Then
            ' Don't check on Mac due to security settings
        #Else
            If Not FileExists(ThisWorkbook.Path & file_path) Then
                MsgBox "Data file not found.", vbExclamation
                Exit Sub
            End If
        #End If
    End If
    
    ' Check for errors in dataset if not an external file (predictive add-in only)
    If Left(frm_pred.lbl_training_data.tag, 5) <> "File:" And f_name = "predictive_addin" Then
        Dim y_var As String, data_range As Range, output_col As Integer
        y_var = Trim(Split(frm_pred.txt_formula.Text, "~")(0))
        Set data_range = Range(remove_workbook_from_range(frm_pred.lbl_training_data.tag))
        If Not (data_range(1, 1) = "ROW" And data_range(1, 2) = "COLUMN" And data_range(1, 3) = "VALUE") Then
            Dim i As Integer, c As Integer
            For i = 1 To data_range.Columns.count
                If data_range(1, i) = y_var Then
                    c = c + 1
                    output_col = i
                End If
            Next i
            If c = 0 Then
                MsgBox "Output variable not found in dataset."
                Exit Sub
            End If
            ' Check for non-numerical values in output column
            Dim j As Integer
            For j = 2 To data_range.Rows.count
                If Not IsNumeric(Trim(data_range(j, output_col))) Then
                    MsgBox "Output variable values must be numeric."
                    Exit Sub
                End If
            Next j
            ' Check for non-numerical values in input variable columns if using dot formula
            If InStr(frm_pred.txt_formula.Text, ".") Then
                Dim k As Integer, col As Integer
                For col = 1 To data_range.Columns.count
                    If data_range(1, col) = y_var Then
                        ' Do nothing, checked above
                    Else
                        For k = 2 To 41
                            If Not IsNumeric(Trim(data_range(k, col))) Then
                                MsgBox "Non-numeric variables are not supported when using a dot formula."
                                Exit Sub
                            End If
                        Next k
                    End If
                Next col
            End If
        End If
    End If
    
    ' Generate a run ID
    run_id True, Asc(Mid(email_provided, 1, 1)) + Asc(Mid(email_provided, 2, 1))
    
    'Log the run on the server
    On Error Resume Next
    
    #If Mac Then
        RunPython "import requests; requests.post(url = 'http://guetta.org/addin/run.php'," & _
                        "data = {'run_id':'" & run_id() & "', " & _
                        "        'version':'" & addin_version() & "', " & _
                        "        'email':'" & email_provided & "', " & _
                        "        'platofrm':'mac'}, timeout = 10)"
    #Else
        windows_curl "http://guetta.org/addin/run.php?run_id=" & run_id() & _
                            "&version=" & addin_version() & _
                            "&email=" & email_provided & _
                            "&platform=windows"
    #End If
    
    On Error GoTo 0
    
    ' Disable all buttons
    disable_buttons
    
    ' Set the status cell to "launching"
    ThisWorkbook.Sheets("Add-in").Range(this_status_cell) = "Launching Python"
    DoEvents
        
    ' Log the start time
    start_time = Timer()
    
    ' Clear any blank sheets that might be lingering
    Dim sheet As Worksheet
    Application.DisplayAlerts = False
    For Each sheet In ThisWorkbook.Sheets
        If WorksheetFunction.CountA(sheet.UsedRange) = 0 And sheet.Visible = 0 And sheet.Name <> "xlwings.conf" Then
            sheet.Delete
        End If
    Next sheet
    Application.DisplayAlerts = True
    
    ' Create a new sheet to output the results - save the sheet's name
    Dim new_sheet_name As String
    new_sheet_name = Sheets.Add().Name
    ActiveWindow.DisplayGridlines = False
    ThisWorkbook.Sheets.Item(new_sheet_name).Visible = xlSheetHidden
    
    On Error Resume Next
    ThisWorkbook.Names.Item("sheet_in_progress").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add "sheet_in_progress", new_sheet_name, False
    
    ' Save the calculation mode and set it to manual so we're not updating while
    ' python is running
    ThisWorkbook.Names.Add "calc_mode", Application.Calculation, False
    Application.Calculation = xlCalculationManual
    
    ' Make sure the configuration is up to date
    update_conf True
    
    ' Figure out whether we're using the UDF
    Dim udf_server As String
    udf_server = "False"
    #If Mac Then
        DoEvents
    #Else
        With Sheets("xlwings.conf")
            If Sheets("Add-in").CheckBoxes("chk_server").Value = 1 Then
                .Range("A1").End(xlDown).Offset(1, 0).Value = "Use UDF Server"
                .Range("B1").End(xlDown).Offset(1, 0).Value = "TRUE"
            End If
            
            If Sheets("Add-in").CheckBoxes("chk_server").Value = 1 Or Sheets("Add-in").CheckBoxes("chk_foreground").Value = 1 Then
                .Range("A1").End(xlDown).Offset(1, 0).Value = "Show Console"
                .Range("B1").End(xlDown).Offset(1, 0).Value = "TRUE"
            End If
        End With
    #End If
    
    ' Run
    Dim code_text As String

    code_text = "import xlwings as xw; import types; mod = types.ModuleType('mod'); " & _
            "xw.Book.caller().sheets('Add-in').range('" & this_status_cell & "').value = 'Python launched; loading packages'; " & _
            "code_sheet = xw.Book.caller().sheets('code_text'); " & _
            "code_range = 'A2:A' + str(int(code_sheet.range('A1').value+1)); " & _
            "exec('\n'.join([str(i) if i is not None else '' for i in code_sheet.range(code_range).value]), mod.__dict__); " & _
            "mod.run_addin('" & f_name & "','" & new_sheet_name & "', " & udf_server & ")"
    
    Dim run_python_result As String
    run_python_result = RunPython(code_text)
    
    If run_python_result <> "" Then
        format_sheet
    End If
End Sub

' ====================
' = Format the sheet =
' ====================

Sub format_sheet()
    ' This function will be called from Python when the add-in is done running;
    ' it will format the sheet and reset the add-in

    ' Begin by resetting the add-in
    ' -----------------------------
    
    ' Save the run ID for the VBA success function
    Dim this_run_id As String
    this_run_id = "unknown"
    
    ' Re-enable the buttons
    enable_buttons
    
    ' Clear status cells
    ThisWorkbook.Sheets("Add-in").Range(STATUS_CELL).Value = ""
    ThisWorkbook.Sheets("Add-in").Range(STATUS_CELL_TEXT).Value = ""

    ' Get the name of the sheet to edit. If we have a run_id, we launched
    ' from Excel and we need to get the sheet name from Excel. Otherwise,
    ' use the standard sheet name
    Dim new_sheet_name As String
    
    If run_id() <> "" Then
        ' Remove the run_id after saving it
        this_run_id = run_id()
        Sheets("code_text").Range("B1").Value = ""
        
        'Get a sheet name if it exists
        On Error Resume Next
            new_sheet_name = ThisWorkbook.Names.Item("sheet_in_progress").Value
            new_sheet_name = Mid(new_sheet_name, 3, Len(new_sheet_name) - 3)
        On Error GoTo 0
        
    Else
        new_sheet_name = "out_sheet_from_python"
    End If
    
    ' Get the sheet itself - if it doesn't exist, leave
    On Error GoTo end_format_sheet
        Dim sht As Worksheet
        Set sht = ThisWorkbook.Sheets(new_sheet_name)
    On Error GoTo 0

    ' If we have no output for some reason, delete the sheet
    If WorksheetFunction.CountA(ThisWorkbook.Sheets(new_sheet_name).UsedRange) = 0 Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(new_sheet_name).Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' Format the results
    ' ==================
    ' The top rows will contain cells that need special formating
    ' treatment.
    Dim font_medium, font_large, bottom_thick, top_thin, italics, align_center, align_right, bold, expand, courier, align_left, graph_formatting, number_format
    Dim num_cols As Integer
    Dim num_rows As Integer
    Dim i As Integer
    
    With ThisWorkbook.Sheets(new_sheet_name)
        font_medium = .Range("A1").Value
        font_large = .Range("A2").Value
        bottom_thick = .Range("A3").Value
        top_thin = .Range("A4").Value
        italics = .Range("A5").Value
        bold = .Range("A6").Value
        align_center = .Range("A7").Value
        align_right = .Range("A8").Value
        expand = .Range("A9").Value
        courier = .Range("A10").Value
        align_left = .Range("A11").Value
        number_format = .Range("A12").Value
        num_cols = .Range("A13").Value
        num_rows = .Range("A14").Value
        
        graph_formatting = .Range("A15").Value
        
        ' Delete these rows
        .Rows("1:15").Delete
    
        If font_medium <> "" Then .Range(font_medium).Font.Size = 16
        If font_large <> "" Then .Range(font_large).Font.Size = 22
        
        If bottom_thick <> "" Then .Range(bottom_thick).Borders(xlEdgeBottom).LineStyle = xlContinuous
        If bottom_thick <> "" Then .Range(bottom_thick).Borders(xlEdgeBottom).Weight = xlMedium
            
        If top_thin <> "" Then .Range(top_thin).Borders(xlEdgeTop).LineStyle = xlContinuous
        If top_thin <> "" Then .Range(top_thin).Borders(xlEdgeTop).Weight = xlThin
        
        If italics <> "" Then .Range(italics).Font.Italic = True
        
        If bold <> "" Then .Range(bold).Font.bold = True
        
        If align_center <> "" Then .Range(align_center).HorizontalAlignment = xlCenter
        
        If align_right <> "" Then .Range(align_right).HorizontalAlignment = xlRight
        
        If number_format <> "" Then .Range(number_format).NumberFormat = "0.000"
        
        If courier <> "" Then .Range(courier).Font.Name = "Courier New"
        
        If align_left <> "" Then .Range(align_left).HorizontalAlignment = xlLeft
        
        If expand <> "" Then .Range(expand).Columns.AutoFit
    
        For i = 1 To num_cols
            If .Columns(i).ColumnWidth > 20 Then .Columns(i).ColumnWidth = 20
        Next i
    
        .Visible = xlSheetVisible
        
        .Activate
        
        If graph_formatting <> "" Then
        
            graph_formatting = Split(graph_formatting, "|")
            
            For i = 0 To UBound(graph_formatting)
                Dim this_instruction
                this_instruction = Split(graph_formatting(i), ",")
                
                Dim inst_2 As Double
                Dim inst_3 As Double
                inst_2 = localize_number(this_instruction(2))
                inst_3 = localize_number(this_instruction(3))
                
                With .Shapes.Range(this_instruction(0))
                    .Top = localize_number(Range(this_instruction(1)).Top)
                    .Left = localize_number(Range(this_instruction(1)).Left)
                    .Width = localize_number(Range(this_instruction(1), Range(this_instruction(1)).Offset(GRAPH_LINE_PER_INCH * inst_2 - 1)).Height)
                    .Height = localize_number(Range(this_instruction(1), Range(this_instruction(1)).Offset(GRAPH_LINE_PER_INCH * inst_2 - 1)).Height) * inst_3 / inst_2
                End With
            Next i
        
        End If
    
    End With
    
    ' Output the overhead time
    With ThisWorkbook.Sheets(new_sheet_name)
        If .Range("A1").Value <> "Text Add-in Error Report" And .Range("A1").Value <> "Add-in Error Report" Then
            .Range("D" & num_rows).Value = Round(Timer - start_time - localize_number(.Range("D" & num_rows).Value), 2)
        End If
    End With
    
    ' Reset the calculation mode
    ' Application.Calculation = Mid(ThisWorkbook.Names.Item("calc_mode"), 2)
    Application.Calculation = xlCalculationAutomatic
        
    ' Log a success
    On Error Resume Next
    #If Mac Then
        RunPython "import requests; requests.post(url = 'http://guetta.org/addin/success.php'," & _
                        "data = {'run_id':'" & this_run_id & "'}, timeout = 10)"
    #Else
        windows_curl "http://guetta.org/addin/success.php?run_id=" & this_run_id
    #End If
    On Error GoTo 0
    
end_format_sheet:
    ' Clean up
    Set sht = Nothing
End Sub
