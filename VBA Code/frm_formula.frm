Dim autocomplete As Boolean

Private Sub cmd_save_Click()
    frm_pred.txt_formula.Text = txt_formula_edit.Text
    Unload Me
End Sub

Private Sub txt_formula_edit_Change()
    On Error GoTo unknown_err
    
    If autocomplete Then
        ' Find the variable stub
        Dim backlook_done As Boolean
        Dim cur_char As String
        Dim var_stub As String
        Dim var_match As String
        Dim start_sel_pos As Integer
        Dim i As Integer
        
        i = txt_formula_edit.SelStart + 1
        While (Not backlook_done) And (i > 1)
            i = i - 1
            cur_char = Mid(txt_formula_edit.Text, i, 1)
            If cur_char = "(" Or cur_char = " " Or cur_char = "+" Or cur_char = ")" _
                    Or cur_char = "~" Or cur_char = "*" Or i <= 1 Then backlook_done = True
        Wend
        If i > 1 Then i = i + 1
        
        If i <= txt_formula_edit.SelStart Then
            var_stub = Mid(txt_formula_edit.Text, i, txt_formula_edit.SelStart - i + 1)
            var_match = variable_valid(var_stub, True)
            
            If var_match <> "" Then
                populate_list var_stub
                autocomplete = False
                
                start_sel_pos = txt_formula_edit.SelStart
                txt_formula_edit.Text = Mid(txt_formula_edit.Text, 1, txt_formula_edit.SelStart) & _
                                          Mid(var_match, Len(var_stub) + 1) & _
                                          Mid(txt_formula_edit.Text, txt_formula_edit.SelStart + 1)
                txt_formula_edit.SelStart = start_sel_pos
                txt_formula_edit.SelLength = Len(var_match) - Len(var_stub)
                
                autocomplete = True
            Else
                populate_list
            End If
        Else
            populate_list
        End If
    End If
    
    Dim error_message As String
    error_message = validate_formula(txt_formula_edit.Text)
    
    GoTo after_err
unknown_err:
    error_message = "Unknown error. Note that this might be a bug in the error-checking code rather than in your formula, so feel free to try and run it anyway and see if it works!"
after_err:
    
    If error_message <> "" Then
        txt_formula_edit.BackColor = RED
        txt_error.Text = error_message
    Else
        txt_formula_edit.BackColor = WHITE
        txt_error.Text = ""
    End If
    
End Sub

Private Sub txt_formula_edit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 8 Then
        autocomplete = False
        populate_list
    Else
        autocomplete = True
    End If
    
    If KeyCode = 9 Then
        KeyCode = 39
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Populate the variable names
    populate_list
    
    ' Initialize the formula and check it
    txt_formula_edit.Text = frm_pred.txt_formula.Text
    
    Dim error_message As String
    error_message = validate_formula(txt_formula_edit.Text)
    
    If error_message <> "" Then
        txt_formula_edit.BackColor = RED
        txt_error.Text = error_message
    Else
        txt_formula_edit.BackColor = WHITE
        txt_error.Text = ""
    End If
    
    ' Turn autocomplete on
    autocomplete = True

    ' Put the window in the ocrrect place
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub populate_list(Optional stub As String = "")
    lst_variables.Clear
    
    If mac_file Then
        lst_variables.AddItem "Mac security settings"
        lst_variables.AddItem "do not allow Excel to"
        lst_variables.AddItem "access the column names"
        lst_variables.AddItem "from a file. You will"
        lst_variables.AddItem "need to open the file,"
        lst_variables.AddItem "look at the column"
        lst_variables.AddItem "names there, and type"
        lst_variables.AddItem "them manually into the"
        lst_variables.AddItem "formula."
    ElseIf invalid_file Then
        lst_variables.AddItem "I've encoutered an error"
        lst_variables.AddItem "trying to find the"
        lst_variables.AddItem "columns in the file you"
        lst_variables.AddItem "selected. You'll need"
        lst_variables.AddItem "type columns here"
        lst_variables.AddItem "manually instead."
    ElseIf excel_file Then
        lst_variables.AddItem "VBA cannot find column"
        lst_variables.AddItem "names for non-CSV files."
        lst_variables.AddItem "You will need to open"
        lst_variables.AddItem "the file, look at the"
        lst_variables.AddItem "column names there, and"
        lst_variables.AddItem "type them manually into"
        lst_variables.AddItem "the formula."
    ElseIf too_many_cols Then
        lst_variables.AddItem "Your data has a lot of"
        lst_variables.AddItem "columns! For the sake"
        lst_variables.AddItem "of not crashing your"
        lst_variables.AddItem "computer, I won't list"
        lst_variables.AddItem "them all here. You can"
        lst_variables.AddItem "manually type them into"
        lst_variables.AddItem "the formula."
    Else
        Dim i As Integer
        For i = 1 To vars.count
            If stub = Mid(vars.Item(i), 1, Len(stub)) Then
                lst_variables.AddItem vars.Item(i)
            End If
        Next i
    End If
End Sub