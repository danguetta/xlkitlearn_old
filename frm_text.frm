Option Explicit

Private Sub chk_sparse_output_Click()
    If chk_sparse_output.Value = True Then
        max_features.Max = 10000
    Else
        max_features.Max = 1000
    End If
End Sub

Private Sub cmd_save_Click()
    ' Save the settings to the worksheet
    
    dict_utils CURRENT_TEXT_SETTINGS, "source_data", txt_source_data.Text
    
    dict_utils CURRENT_TEXT_SETTINGS, "max_df", txt_max_df.Text
    dict_utils CURRENT_TEXT_SETTINGS, "min_df", txt_min_df.Text
    dict_utils CURRENT_TEXT_SETTINGS, "max_features", max_features.Value
    dict_utils CURRENT_TEXT_SETTINGS, "stop_words", chk_stop_words.Value
    dict_utils CURRENT_TEXT_SETTINGS, "tf_idf", chk_tf_idf.Value
    dict_utils CURRENT_TEXT_SETTINGS, "bigrams", chk_bigrams.Value
    dict_utils CURRENT_TEXT_SETTINGS, "stem", chk_stem.Value
    
    dict_utils CURRENT_TEXT_SETTINGS, "lda_topics", txt_lda_topics.Text
    dict_utils CURRENT_TEXT_SETTINGS, "eval_perc", txt_eval_perc.Text
    
    dict_utils CURRENT_TEXT_SETTINGS, "seed", txt_seed.Text
    
    dict_utils CURRENT_TEXT_SETTINGS, "output_code", chk_output_code.Value
    
    dict_utils CURRENT_TEXT_SETTINGS, "max_lda_iter", txt_max_lda_iter.Text
    
    dict_utils CURRENT_TEXT_SETTINGS, "sparse_output", chk_sparse_output.Value
    
    Unload Me
End Sub

Private Sub max_features_Change()
    lbl_max_features.caption = max_features.Value
End Sub

Private Sub opt_features_Click()
    txt_eval_perc.Enabled = True
    chk_sparse_output.Enabled = True
    max_features.Max = 1000
    
    With txt_lda_topics
        .Enabled = False
        .Value = ""
    End With
    With txt_max_lda_iter
        .Enabled = False
        .Value = ""
    End With
End Sub

Private Sub opt_lda_Click()
    With txt_eval_perc
        .Enabled = False
        .Text = ""
    End With
    chk_sparse_output.Enabled = False
    chk_sparse_output.Value = False
    
    txt_lda_topics.Enabled = True
    txt_max_lda_iter.Enabled = True
    max_features.Max = 10000
End Sub


Private Sub txt_eval_perc_Change()
    validate_parameters
End Sub

Private Sub txt_lda_topics_Change()
    validate_parameters
End Sub

Private Sub txt_max_df_Change()
    validate_parameters
End Sub

Private Sub txt_max_lda_iter_Change()
    validate_parameters
End Sub

Private Sub txt_min_df_Change()
    validate_parameters
End Sub

Private Sub txt_seed_Change()
    validate_parameters
End Sub

Private Sub txt_source_data_Change()
    validate_parameters
End Sub

Private Sub UserForm_Initialize()
   
    ' Sync the dicts
    Range(CURRENT_TEXT_SETTINGS).Value = sync_dicts(Range(CURRENT_TEXT_SETTINGS).Value, BLANK_TEXT_SETTINGS)
        
    ' Populate the form
    txt_source_data.Text = dict_utils(CURRENT_TEXT_SETTINGS, "source_data")
    
    txt_max_df.Text = dict_utils(CURRENT_TEXT_SETTINGS, "max_df")
    txt_min_df.Text = dict_utils(CURRENT_TEXT_SETTINGS, "min_df")
    max_features.Value = dict_utils(CURRENT_TEXT_SETTINGS, "max_features")
    max_features_Change
    
    chk_stop_words.Value = dict_utils(CURRENT_TEXT_SETTINGS, "stop_words")
    chk_tf_idf.Value = dict_utils(CURRENT_TEXT_SETTINGS, "tf_idf")
    chk_bigrams.Value = dict_utils(CURRENT_TEXT_SETTINGS, "bigrams")
    chk_stem.Value = dict_utils(CURRENT_TEXT_SETTINGS, "stem")
    
    txt_lda_topics.Text = dict_utils(CURRENT_TEXT_SETTINGS, "lda_topics")
    If txt_lda_topics.Text <> "" Then
        opt_lda.Value = True
        opt_lda_Click
    End If
    
    txt_eval_perc.Text = dict_utils(CURRENT_TEXT_SETTINGS, "eval_perc")
    If txt_eval_perc.Text <> "" Then
        opt_features.Value = True
        opt_features_Click
    End If
    
    txt_seed.Text = dict_utils(CURRENT_TEXT_SETTINGS, "seed")
    
    chk_output_code.Value = dict_utils(CURRENT_TEXT_SETTINGS, "output_code")
    txt_max_lda_iter.Text = dict_utils(CURRENT_TEXT_SETTINGS, "max_lda_iter")
    
    chk_sparse_output.Value = dict_utils(CURRENT_TEXT_SETTINGS, "sparse_output")
    chk_sparse_output_Click
    
    ' Put the window in the ocrrect place
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


Private Sub validate_parameters()
    ' This function will validate parameters and indicate any mistakes
    Const NON_NUMERIC = "This parameter needs to be a number; please correct."
    Const FILE_NOT_EXIST = "This file might not exist. Feel free to run the add-in anyway and see if it works; sometimes Excel has trouble identifying files."
    Const DATA_NEEDED = "Please provide a file with text data."
    Const FREQ_OUT_OF_RANGE = "This parameter needs to be a number between 0 and 1."
    Const PERC_OUT_OF_RANGE = "This parameter is a percentage and needs to be between 0 and 100."
    Const MIN_MAX_FLIPPED = "The maximum must be greater than the minimum."
    Const TOPICS_BELOW_TWO = "This parameter must be greater than 2."
    Const SEED_NOT_GIVEN = "A random seed is required when splitting the data into test/train."
    
    ' Clear all the entries
    txt_source_data.BackColor = WHITE
    txt_min_df.BackColor = WHITE
    txt_max_df.BackColor = WHITE
    txt_eval_perc.BackColor = WHITE
    txt_lda_topics.BackColor = WHITE
    txt_max_lda_iter.BackColor = WHITE
    txt_seed.BackColor = WHITE
    
    txt_source_data.ControlTipText = ""
    txt_min_df.ControlTipText = ""
    txt_max_df.ControlTipText = ""
    txt_eval_perc.ControlTipText = ""
    txt_lda_topics.ControlTipText = ""
    txt_max_lda_iter.ControlTipText = ""
    txt_seed.ControlTipText = ""
    
    ' Check for numeric entries
    If (Not IsNumeric(txt_min_df.Text)) And txt_min_df.Text <> "" Then
        txt_min_df.BackColor = RED
        txt_min_df.ControlTipText = NON_NUMERIC
    ElseIf (txt_min_df.Text <> "") And (Trim(txt_min_df.Text) < 0 Or Trim(txt_min_df.Text) > 1) Then
        txt_min_df.BackColor = RED
        txt_min_df.ControlTipText = FREQ_OUT_OF_RANGE
    End If
    
    If (Not IsNumeric(txt_max_df.Text)) And txt_max_df.Text <> "" Then
        txt_max_df.BackColor = RED
        txt_max_df.ControlTipText = NON_NUMERIC
    ElseIf (txt_max_df.Text <> "") And (Trim(txt_max_df.Text) < 0 Or Trim(txt_max_df.Text) > 1) Then
        txt_max_df.BackColor = RED
        txt_max_df.ControlTipText = FREQ_OUT_OF_RANGE
    ElseIf (txt_min_df.Text <> "") And (txt_max_df.Text <> "") Then
        If Trim(txt_min_df.Text) > Trim(txt_max_df.Text) Then
            txt_max_df.BackColor = RED
            txt_max_df.ControlTipText = MIN_MAX_FLIPPED
        End If
    End If
    
    If (Not IsNumeric(txt_eval_perc.Text)) And txt_eval_perc.Text <> "" Then
        txt_eval_perc.BackColor = RED
        txt_eval_perc.ControlTipText = NON_NUMERIC
    ElseIf txt_eval_perc.Text <> "" And (Trim(txt_eval_perc.Text) < 0 Or Trim(txt_eval_perc.Text) > 100) Then
        txt_eval_perc.BackColor = RED
        txt_eval_perc.ControlTipText = PERC_OUT_OF_RANGE
    End If
    
    If (Not IsNumeric(txt_lda_topics.Text)) And txt_lda_topics.Text <> "" Then
        txt_lda_topics.BackColor = RED
        txt_lda_topics.ControlTipText = NON_NUMERIC
    ElseIf txt_lda_topics.Text <> "" And (Trim(txt_lda_topics.Text) < 2) Then
        txt_lda_topics.BackColor = RED
        txt_lda_topics.ControlTipText = TOPICS_BELOW_TWO
    End If
    
    If (Not IsNumeric(txt_max_lda_iter.Text)) And txt_max_lda_iter.Text <> "" Then
        txt_max_lda_iter.BackColor = RED
        txt_max_lda_iter.ControlTipText = NON_NUMERIC
    End If
    
    If (Not IsNumeric(txt_seed.Text)) And txt_seed.Text <> "" Then
        txt_seed.BackColor = RED
        txt_seed.ControlTipText = NON_NUMERIC
    ElseIf txt_eval_perc.Text <> "" And txt_seed.Text = "" Then
        txt_seed.BackColor = RED
        txt_seed.ControlTipText = SEED_NOT_GIVEN
    End If
    
    ' Ensure the file exists
    If txt_source_data.Text = "" Then
        txt_source_data.BackColor = RED
        txt_source_data.ControlTipText = DATA_NEEDED
    Else:
        #If Mac Then
            DoEvents
        #Else
            If Not FileExists(get_full_file_name(txt_source_data.Text)) Then
                txt_source_data.BackColor = RED
                txt_source_data.ControlTipText = FILE_NOT_EXIST
            End If
        #End If
    End If
End Sub


