Option Explicit

Private Sub cmb_model_Change()
    ' Get the list of parameters for this model
    Dim param_list
    param_list = Split(dict_utils(MODELS, cmb_model.Value), ",")
    
    txt_param1.Text = ""
    txt_param2.Text = ""
    txt_param3.Text = ""
    txt_param1.Visible = False
    txt_param2.Visible = False
    txt_param3.Visible = False
    lbl_param1.Visible = False
    lbl_param2.Visible = False
    lbl_param3.Visible = False
    
    If UBound(param_list) >= 0 Then
        lbl_param1.caption = param_list(0)
        lbl_param1.Visible = True
        txt_param1.Visible = True
    End If
    
    If UBound(param_list) >= 1 Then
        lbl_param2.caption = param_list(1)
        lbl_param2.Visible = True
        txt_param2.Visible = True
    End If
    
    If UBound(param_list) >= 2 Then
        lbl_param3.caption = param_list(2)
        lbl_param3.Visible = True
        txt_param3.Visible = True
    End If
    
    validate_parameters
End Sub


Private Sub cmd_save_Click()
    Application.Calculation = xlCalculationManual
    
    ' Save the settings to the worksheet
    
    dict_utils CURRENT_SETTINGS, "model", cmb_model.Text
    dict_utils CURRENT_SETTINGS, "formula", txt_formula.Text
    dict_utils CURRENT_SETTINGS, "output_model", chk_output_model.Value
    dict_utils CURRENT_SETTINGS, "output_code", chk_output_code.Value
    
    dict_utils CURRENT_SETTINGS, "param1", txt_param1.Text
    dict_utils CURRENT_SETTINGS, "param2", txt_param2.Text
    dict_utils CURRENT_SETTINGS, "param3", txt_param3.Text
    
    dict_utils CURRENT_SETTINGS, "training_data", lbl_training_data.tag
    
    dict_utils CURRENT_SETTINGS, "seed", txt_seed.Text
    
    dict_utils CURRENT_SETTINGS, "K", txt_K.Text
    
    ' Change "False" to chk_ts_data.value to use parameter
    dict_utils CURRENT_SETTINGS, "ts_data", "False"
    
    dict_utils CURRENT_SETTINGS, "evaluation_perc", txt_evaluation_perc.Text
    dict_utils CURRENT_SETTINGS, "evaluation_data", lbl_evaluation_data.tag
    dict_utils CURRENT_SETTINGS, "output_evaluation_details", chk_output_evaluation_details.Value
    
    dict_utils CURRENT_SETTINGS, "prediction_data", lbl_prediction_data.tag
    
    Application.Calculation = xlCalculationAutomatic
    
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    frm_formula.Show False
End Sub

Private Sub lbl_evaluation_data_Click()
    Me.Hide
    
    lbl_evaluation_data.tag = get_range("Select the evaluation range or file. EITHER with headers matching training data in the first row (in any order) " & _
                                            "OR without headers but containing the same number of columns as the training data, " & _
                                            "in the same order.", lbl_prediction_data.tag, True, True)
                                            
    If lbl_evaluation_data.tag = "" Then
        lbl_evaluation_data.caption = BLANK_RANGE_SELECTOR
    ElseIf Left(lbl_evaluation_data.tag, 6) = "File: " Then
        lbl_evaluation_data.caption = lbl_evaluation_data.tag
        clear_vars
    Else
        lbl_evaluation_data.caption = lbl_evaluation_data.tag
    End If
    
    Me.Show False
    validate_parameters
End Sub

Private Sub lbl_prediction_data_Click()
    Me.Hide
                                  
    lbl_prediction_data.tag = get_range("Select the prediction range or file. EITHER with headers matching training data in the first row (in any order) " & _
                                            "OR without headers but containing the same number of columns as the training data, " & _
                                            "in the same order.", lbl_prediction_data.tag, True, True)
               
    If lbl_prediction_data.tag = "" Then
        lbl_prediction_data.caption = BLANK_RANGE_SELECTOR
    ElseIf Left(lbl_prediction_data.tag, 6) = "File: " Then
        lbl_prediction_data.caption = lbl_prediction_data.tag
        clear_vars
    Else
        lbl_prediction_data.caption = lbl_prediction_data.tag
    End If
    
    Me.Show False
    validate_parameters
End Sub

Private Sub lbl_training_data_Click()
    Me.Hide
    lbl_training_data.tag = get_range("Select a range for the training data." & vbCrLf & vbCrLf & _
                                        "The first row of the table you select should contain variable names." & _
                                        vbCrLf & vbCrLf & "You can also enter a file name, provided it exists in the " & _
                                        "same directory as this spreadsheet.", _
                                        lbl_training_data.tag, True, True)
    
    If lbl_training_data.tag = "" Then
        lbl_training_data.caption = BLANK_RANGE_SELECTOR
        clear_vars
    Else
        lbl_training_data.caption = lbl_training_data.tag
        update_vars
    End If
    
    Me.Show False
    validate_parameters
End Sub

Private Sub lbl_training_errors_Click()
    frm_data_errors.Show False
End Sub

Private Sub opt_no_eval_Click()
    With txt_evaluation_perc
        .Text = ""
        .Enabled = False
    End With
    
    With lbl_evaluation_data
        .tag = ""
        .caption = BLANK_RANGE_SELECTOR
        .Enabled = False
    End With
    validate_parameters
End Sub

Private Sub opt_perc_eval_Click()
    With txt_evaluation_perc
        .Enabled = True
    End With
    
    With lbl_evaluation_data
        .tag = ""
        .caption = BLANK_RANGE_SELECTOR
        .Enabled = False
    End With
    validate_parameters
End Sub

Private Sub opt_specific_eval_Click()
    With txt_evaluation_perc
        .Text = ""
        .Enabled = False
    End With
    
    With lbl_evaluation_data
        .Enabled = True
    End With
    validate_parameters
End Sub

Private Sub txt_evaluation_perc_Change()
    validate_parameters
End Sub

Private Sub txt_formula_Change()
    validate_parameters
End Sub

Private Sub txt_formula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    #If Mac Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 0
            txt_formula.Text = GetClipBoardText()
        End If
    #End If
End Sub


Private Sub txt_K_Change()
    validate_parameters
End Sub

Private Sub txt_param1_Change()
    validate_parameters
End Sub

Private Sub txt_param1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    #If Mac Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 0
            txt_param1.Text = GetClipBoardText()
        End If
    #End If
End Sub

Private Sub txt_param2_Change()
    validate_parameters
End Sub

Private Sub txt_param2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    #If Mac Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 0
            txt_param2.Text = GetClipBoardText()
        End If
    #End If
End Sub

Private Sub txt_param3_Change()
    validate_parameters
End Sub

Private Sub txt_param3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    #If Mac Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 0
            txt_param3.Text = GetClipBoardText()
        End If
    #End If
End Sub

Private Sub txt_seed_Change()
    validate_parameters
End Sub

Private Sub UserForm_Initialize()
    'Fix vagrant texboxes
    txt_K.Left = 170
    txt_evaluation_perc.Left = 250
    
    ' Populate the models combo box
    Dim i As Integer
    Dim all_models
    all_models = dict_utils(MODELS)
    
    For i = 0 To UBound(all_models)
        cmb_model.AddItem (all_models(i))
    Next i
    
    ' Sync the dicts
    Range(CURRENT_SETTINGS).Value = sync_dicts(Range(CURRENT_SETTINGS).Value, BLANK_SETTINGS)
        
    ' Populate the form
    cmb_model.Text = dict_utils(CURRENT_SETTINGS, "model")
    txt_formula.Text = dict_utils(CURRENT_SETTINGS, "formula")
    chk_output_model.Value = dict_utils(CURRENT_SETTINGS, "output_model")
    chk_output_code.Value = dict_utils(CURRENT_SETTINGS, "output_code")
    
    txt_param1.Text = dict_utils(CURRENT_SETTINGS, "param1")
    txt_param2.Text = dict_utils(CURRENT_SETTINGS, "param2")
    txt_param3.Text = dict_utils(CURRENT_SETTINGS, "param3")
    
    clear_vars
    lbl_training_data.tag = dict_utils(CURRENT_SETTINGS, "training_data")
    If lbl_training_data.tag <> "" Then
        lbl_training_data.caption = lbl_training_data.tag
        If check_valid_range(lbl_training_data.caption) Or Left(lbl_training_data.tag, 5) = "File:" Then
            update_vars
        End If
    Else
        lbl_training_data.caption = BLANK_RANGE_SELECTOR
    End If
    
    txt_seed.Text = dict_utils(CURRENT_SETTINGS, "seed")
    
    txt_K.Text = dict_utils(CURRENT_SETTINGS, "K")
    
    ' Uncomment when time series dat is introduced
    ' chk_ts_data.value = dict_utils(CURRENT_SETTINGS, "ts_data")
    
    txt_evaluation_perc.Text = dict_utils(CURRENT_SETTINGS, "evaluation_perc")
    If txt_evaluation_perc.Text <> "" Then
        opt_perc_eval.Value = True
        opt_perc_eval_Click
    End If
    
    lbl_evaluation_data.tag = dict_utils(CURRENT_SETTINGS, "evaluation_data")
    If lbl_evaluation_data.tag <> "" Then
        lbl_evaluation_data.caption = lbl_evaluation_data.tag
        opt_specific_eval.Value = True
        opt_specific_eval_Click
    Else
        lbl_evaluation_data.caption = BLANK_RANGE_SELECTOR
    End If
    chk_output_evaluation_details.Value = dict_utils(CURRENT_SETTINGS, "output_evaluation_details")
        
    lbl_prediction_data.tag = dict_utils(CURRENT_SETTINGS, "prediction_data")
    If lbl_prediction_data.tag <> "" Then
        lbl_prediction_data.caption = lbl_prediction_data.tag
    Else
        lbl_prediction_data.caption = BLANK_RANGE_SELECTOR
    End If
    
    ' Add tooltips for paste in mac
    #If Mac Then
        txt_formula.ControlTipText = "To replace the text in this textbox with the text in your clipboard, press Ctrl + V"
    #End If
    
    ' Validate parameters
    validate_parameters
    
    ' Put the window in the ocrrect place
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Public Function isNumericList(list_string As String) As Boolean
    Dim i As Integer
    Dim split_list As Variant
    
    split_list = Split(list_string, "&")
    
    For i = 0 To UBound(split_list)
        If Not IsNumeric(Trim(split_list(i))) Then
            isNumericList = False
            Exit Function
        End If
    Next i
    
    isNumericList = True
End Function

Public Function isWeightList(list_string As String) As Boolean
    Dim i As Integer
    Dim split_list As Variant
    
    split_list = Split(list_string, "&")
    
    For i = 0 To UBound(split_list)
        If Not (Trim(split_list(i)) = "u" Or Trim(split_list(i)) = "d" Or Trim(split_list(i)) = "uniform" Or Trim(split_list(i)) = "distance") Then
            isWeightList = False
            Exit Function
        End If
    Next i
    
    isWeightList = True

End Function

Public Sub clear_vars()
    While vars.count > 0
        vars.Remove (1)
    Wend
    
    While var_types.count > 0
        var_types.Remove (1)
    Wend
End Sub

Public Sub clear_header_errors()
    While header_errors.count > 0
        header_errors.Remove (1)
    Wend
End Sub

Public Sub update_vars()
    On Error GoTo clear_all
    
    clear_vars
    
    Dim i As Integer
    
    If Left(frm_pred.lbl_training_data.tag, 5) = "File:" Then
        mac_file = False
        excel_file = False
        too_many_cols = False
        invalid_file = False
        
        #If Mac Then
            mac_file = True
            Exit Sub
        #Else
            Dim file_name As String
            file_name = Split(frm_pred.lbl_training_data.tag, ": ")(1)
            If Right(file_name, 4) = ".csv" Then
                If FileExists(get_full_file_name(file_name)) Then
                    Dim data_line
                    Open get_full_file_name(file_name) For Input As #1
                    Line Input #1, data_line
                    Close #1
                    
                    data_line = Split(data_line, ",")
                    
                    If Trim(data_line(0)) = "ROW" And Trim(data_line(1)) = "COLUMN" And Trim(data_line(2)) = "VALUE" Then
                        Exit Sub
                    End If
                    
                    If UBound(data_line) < 50 Then
                        For i = 0 To UBound(data_line)
                            If Mid(data_line(i), 1, 1) = """" And Right(data_line(i), 1) = """" Then
                                vars.Add Mid(data_line(i), 2, Len(data_line(i)) - 2)
                            Else
                                vars.Add data_line(i)
                            End If
                        Next i
                        vars.Add "."
                    Else
                        too_many_cols = True
                    End If
                    
                Else
                    invalid_file = True
                End If
                
                Exit Sub
            Else
                excel_file = True
                Exit Sub
            End If
        #End If
    End If
    
    Dim data_range As Range
    
    Set data_range = Range(remove_workbook_from_range(frm_pred.lbl_training_data.tag))
    
    If data_range(1, 1) = "ROW" And data_range(1, 2) = "COLUMN" And data_range(1, 3) = "VALUE" Then
        ' We have a sparse dataset - column names are blank
    ElseIf data_range.Columns.count < 50 Then
        For i = 1 To data_range.Columns.count
            vars.Add Trim(data_range(1, i))
        Next i
        
        vars.Add "."
        
        ' We have an Excel range with fewer than 50 columns, and a dense datataset. Do some
        ' due diligence and see whether these columns contain any non-numeric data
        For i = 1 To data_range.Columns.count
            Dim this_col As Range
            Dim this_count As Long
            Dim this_countA As Long
            Dim this_len As Long
            
            this_len = data_range.Rows.count - 1
            this_count = WorksheetFunction.count(data_range.Columns(i).Offset(1).Resize(this_len))
            this_countA = WorksheetFunction.CountA(data_range.Columns(i).Offset(1).Resize(this_len))
            
            If this_countA <> this_len Then
                var_types.Add "missing"
            ElseIf this_countA <> this_count Then
                var_types.Add "string"
            Else
                var_types.Add "numeric"
            End If
        Next i
    Else
        too_many_cols = True
    End If
    
    Exit Sub
clear_all:
    clear_vars
End Sub

Private Sub validate_parameters()
    ' This function will validate parameters and indicate any mistakes
    Const NON_NUMERIC = "This parameter needs to be a number; please correct."
    Const PARAM_NEEDED = "This parameter cannot be empty."
    Const DATA_NEEDED = "Training data is needed."
    Const KNN_INVALID = "This parameter must either be 'u', 'd', 'uniform', or 'distance'."
    Const PERC_OUT_OF_RANGE = "This parameter is a percentage and must be between 0 and 100."
    Const BS_GT1_FORMULA = "Best-subset selection can only be done with one formula."
    Const BS_GT10_VARS = "Best-subset selection with more than 10 variables would result in over 1000 competing models." & _
                                            " Consider using something more robust than XLKitLearn."
    Const BS_WITH_INTERCEPT = "XLKitLearn doesn't support best-subset selection with an intercept suppressing term."
    Const HEADERS_DONT_MATCH = "Headers in evaluation or prediction dataset must match headers in training set."
    
    Const SPACE_HEADER = "This column name contains a space"
    Const RESERVED_HEADER = "The name of this column is a 'reserved' Python keyword; try something else"
    Const NUM_START_HEADER = "Column names cannot start with a number"
    Const DUPLICATE_HEADER = "This column name is repeated; column names can only appear once"
    Const INVALID_VAR_CHAR = "Variable names can only contain letters, numbers, and underscores"
                                                                
    Dim UNSUPPORTED_HEADERS As Variant
    UNSUPPORTED_HEADERS = Array("intercept", "and", "as", "assert", "break", "class", "continue", "def", "del", "elif", _
                                            "else", "except", "false", "finally", "for", "from", "global", "if", "import", "in", "is", "lambda", _
                                            "none", "nonlocal", "not", "or", "pass", "raise", "return", "true", "try", "while", "with", "yield")
    
    
    Const FORMULA_WIDTH_LARGE = 198
    Const FORMULA_WIDTH_SMALL = 132
    
    ' Start assuming everything is OK at a global level
    pred_errors = False
    
    ' Clear all the entries
    txt_formula.BackColor = WHITE
    txt_param1.BackColor = WHITE
    txt_param2.BackColor = WHITE
    txt_param3.BackColor = WHITE
    lbl_training_data.BackColor = WHITE
    txt_seed.BackColor = WHITE
    txt_K.BackColor = WHITE
    txt_evaluation_perc.BackColor = WHITE
    lbl_evaluation_data.BackColor = WHITE
    lbl_prediction_data.BackColor = WHITE
    
    txt_formula.ControlTipText = ""
    txt_param1.ControlTipText = ""
    txt_param2.ControlTipText = ""
    txt_param3.ControlTipText = ""
    lbl_training_data.ControlTipText = ""
    txt_seed.ControlTipText = ""
    txt_K.ControlTipText = ""
    txt_evaluation_perc.ControlTipText = ""
    lbl_evaluation_data.ControlTipText = ""
    lbl_prediction_data.ControlTipText = ""
    
    lbl_training_data.Width = FORMULA_WIDTH_LARGE
    lbl_training_errors.Visible = False
    
    ' Determine the number of parameters in th emodel in question
    Dim n_params_required As Integer
    n_params_required = UBound(Split(dict_utils(MODELS, cmb_model.Value), ",")) + 1
    
    ' Check parameters for errors
    If (cmb_model.Value = "Linear/logistic regression") And (LCase(txt_param1.Text) = "bs") Then
        ' If doing best subset, make sure using only one formula, without intercept suppressing term, and with less than 10 vars
        If InStr(Trim(txt_formula.Text), "&") <> 0 Then
            txt_formula.BackColor = RED
            txt_formula.ControlTipText = BS_GT1_FORMULA
            pred_errors = True
        ElseIf InStr(Trim(txt_formula.Text), "-1") <> 0 Then
            txt_formula.BackColor = RED
            txt_formula.ControlTipText = BS_WITH_INTERCEPT
            pred_errors = True
        Else
            If InStr(Trim(txt_formula.Text), ".") = 0 Then
                Dim count As Integer
                ' Find the number of terms separated by a "+" in the formula
                count = UBound(Split(txt_formula.Text, "+")) + 1
                If count >= 10 Then
                    txt_formula.BackColor = RED
                    txt_formula.ControlTipText = BS_GT10_VARS
                    pred_errors = True
                End If
            Else
                If vars.count >= 10 Then
                    txt_formula.BackColor = RED
                    txt_formula.ControlTipText = BS_GT10_VARS
                    pred_errors = True
                End If
            End If
        End If
    ElseIf (Not isNumericList(txt_param1.Text)) Then
        txt_param1.BackColor = RED
        txt_param1.ControlTipText = NON_NUMERIC
        pred_errors = True
    ElseIf Trim(txt_param1.Text) = "" And cmb_model.Value <> "Linear/logistic regression" Then
        txt_param1.BackColor = RED
        txt_param1.ControlTipText = PARAM_NEEDED
        pred_errors = True
    End If
    
    If n_params_required >= 2 Then
        If (Not isNumericList(txt_param2.Text)) And cmb_model.Value <> "K-Nearest Neighbors" Then
            txt_param2.BackColor = RED
            txt_param2.ControlTipText = NON_NUMERIC
            pred_errors = True
        ElseIf Trim(txt_param2.Text) = "" And cmb_model.Value <> "K-Nearest Neighbors" Then
            txt_param2.BackColor = RED
            txt_param2.ControlTipText = PARAM_NEEDED
            pred_errors = True
        End If
    
        If Trim(txt_param2.Text) <> "" And cmb_model.Value = "K-Nearest Neighbors" Then
            If Not (isWeightList(txt_param2.Text)) Then
                txt_param2.BackColor = RED
                txt_param2.ControlTipText = KNN_INVALID
                pred_errors = True
            End If
        End If
    End If
    
    If n_params_required >= 3 Then
        If Not isNumericList(txt_param3.Text) Then
            txt_param3.BackColor = RED
            txt_param3.ControlTipText = NON_NUMERIC
            pred_errors = True
        ElseIf (Trim(txt_param3.Text) = "") And (cmb_model.Value <> "Boosted decision tree") Then
            txt_param3.BackColor = RED
            txt_param3.ControlTipText = PARAM_NEEDED
            pred_errors = True
        End If
    End If

    If Not IsNumeric(txt_seed.Text) And Trim(txt_seed.Text) <> "" Then
        txt_seed.BackColor = RED
        txt_seed.ControlTipText = NON_NUMERIC
        pred_errors = True
    End If

    If Not IsNumeric(txt_K.Text) And Trim(txt_K.Text) <> "" Then
        txt_K.BackColor = RED
        txt_K.ControlTipText = NON_NUMERIC
        pred_errors = True
    ElseIf Trim(txt_K.Text) = "" And ((InStr(1, txt_formula.Text & txt_param1.Text _
                                    & txt_param2.Text & txt_param3.Text, "&") > 0) Or cmb_model.Value = "Boosted decision tree") Then
        txt_K.BackColor = RED
        txt_K.ControlTipText = PARAM_NEEDED
        pred_errors = True
    End If
    
    If Not IsNumeric(txt_evaluation_perc.Text) And Trim(txt_evaluation_perc.Text) <> "" Then
        txt_evaluation_perc.BackColor = RED
        txt_evaluation_perc.ControlTipText = NON_NUMERIC
        pred_errors = True
    ElseIf Trim(txt_evaluation_perc.Text) <> "" And (Trim(txt_evaluation_perc.Text) < 0 Or Trim(txt_evaluation_perc.Text) > 100) Then
        txt_evaluation_perc.BackColor = RED
        txt_evaluation_perc.ControlTipText = PERC_OUT_OF_RANGE
        pred_errors = True
    ElseIf Trim(txt_evaluation_perc.Text) = "" And opt_perc_eval.Value = True Then
        txt_evaluation_perc.BackColor = RED
        txt_evaluation_perc.ControlTipText = PARAM_NEEDED
        pred_errors = True
    End If
    
    If (Not IsNumeric(txt_seed.Text)) And (txt_seed.Text <> "") Then
        txt_seed.BackColor = RED
        txt_seed.ControlTipText = NON_NUMERIC
        pred_errors = True
    End If
    
    ' Check the data
    If lbl_training_data.tag = "" Then
        lbl_training_data.BackColor = RED
        lbl_training_data.ControlTipText = DATA_NEEDED
        pred_errors = True
    Else
        clear_header_errors
        
        ' If on Mac and using external file, vars will be empty and this whole section
        ' will effectively be skipped. "-1" ensures we don't catch the "." variable
        ' name
        
        Dim i As Integer
        For i = 1 To vars.count - 1
            If InStr(1, vars(i), " ") <> 0 Then
                header_errors.Add i & "`" & Replace(vars(i), "`", "") & "`" & SPACE_HEADER
            ElseIf Not IsError(Application.Match(Trim(LCase(vars(i))), UNSUPPORTED_HEADERS, 0)) Then
                header_errors.Add i & "`" & Replace(vars(i), "`", "") & "`" & RESERVED_HEADER
            ElseIf IsNumeric(Left(vars(i), 1)) Then
                header_errors.Add i & "`" & Replace(vars(i), "`", "") & "`" & NUM_START_HEADER
            ElseIf Not valid_var_chars(vars(i)) Then
                header_errors.Add i & "`" & Replace(vars(i), "`", "") & "`" & INVALID_VAR_CHAR
            Else
                ' Count number of times header appears
                Dim j As Integer, c As Integer
                c = 0
                For j = 1 To vars.count
                    If vars(i) = vars(j) Then c = c + 1
                Next j
                
                If c > 1 Then
                    header_errors.Add i & "`" & Replace(vars(i), "`", "") & "`" & DUPLICATE_HEADER
                End If
            End If
        Next i
        
        If header_errors.count > 0 Then
            lbl_training_data.Width = FORMULA_WIDTH_SMALL
            lbl_training_data.BackColor = RED
            
            lbl_training_errors.BackColor = RED
            lbl_training_errors.Visible = True
            
            pred_errors = True
        End If
        
    End If
    
    ' Check that evaluation and prediction data headers match training data
    If Left(lbl_training_data.tag, 5) <> "File:" Then
        If lbl_evaluation_data.tag <> "" And Left(lbl_evaluation_data.tag, 5) <> "File:" Then
            Dim eval_range As Range
            Set eval_range = Range(remove_workbook_from_range(lbl_evaluation_data.tag))
            c = 0
            For i = 1 To eval_range.Columns.count
                For j = 1 To vars.count
                    If eval_range(1, i) = vars(j) Then c = c + 1
                Next j
            Next i
            If c = 0 And (eval_range.Columns.count <> (vars.count - 1)) Then
                lbl_evaluation_data.BackColor = RED
                lbl_evaluation_data.ControlTipText = HEADERS_DONT_MATCH
                pred_errors = True
            End If
        End If
        
        If lbl_prediction_data.tag <> "" And Left(lbl_prediction_data.tag, 5) <> "File:" Then
            Dim pred_range As Range
            Set pred_range = Range(remove_workbook_from_range(lbl_prediction_data.tag))
            c = 0
            For i = 1 To pred_range.Columns.count
                For j = 1 To vars.count
                    If pred_range(1, i) = vars(j) Then c = c + 1
                Next j
            Next i
            If c = 0 And (pred_range.Columns.count <> (vars.count - 1)) Then
                lbl_prediction_data.BackColor = RED
                lbl_prediction_data.ControlTipText = HEADERS_DONT_MATCH
                pred_errors = True
            End If
        End If
    End If
    
    ' Check the formula and data
    Dim error_message As String
    error_message = validate_formula(txt_formula.Text)
    
    If error_message <> "" Then
        txt_formula.BackColor = RED
        txt_formula.ControlTipText = error_message
        pred_errors = True
    End If
    
End Sub


