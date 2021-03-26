Option Explicit

Public vars As New Collection
Public mac_file As Boolean
Public excel_file As Boolean
Public too_many_cols As Boolean
Public invalid_file As Boolean

' Change this to True before releasing
'   (note: if upgrading to a new version of xlwings, make sure show_error
'          in xlwings logs to the server)
' Also deeply hide the sheets
' Also remove the path from the code_text
Public PROD_VERSION As Boolean

Public Const AUTH_KEY = "eMSpn8v3Dc5GoQT7g76NDWfYGg4attWJKg6N3kHt1GdfN0iN321CluEez8tW1trIMcSGdj0sEgkNEuO8NtqHyHah5Qdn05hDZCT"

Public Const MISSING = "XxXxXxXxXxXxX"

Public Const MODELS = "{'Linear/logistic regression'|'Lasso penalty'`'Decision tree'|'Tree depth'`'Boosted decision tree'|'Tree depth,Max trees,Learning rate'`'Random forest'|'Tree depth,Number of trees'}"
Public Const BLANK_SETTINGS = "{'model'|'Linear/logistic regression'`'formula'|''`'param1'|''`'param2'|''`'param3'|''`'training_data'|''`'K'|''`'ts_data'|'False'`'evaluation_perc'|''`'evaluation_data'|''`'prediction_data'|''`'seed'|'123'`'output_model'|'True'`'output_evaluation_details'|'True'`'output_code'|'True'}"
Public Const BLANK_TEXT_SETTINGS = "{'source_data'|''`'max_df'|''`'min_df'|''`'max_features'|'500'`'stop_words'|'True'`'tf_idf'|'False'`'lda_topics'|''`'seed'|'123'`'eval_perc'|'0'`'bigrams'|'False'`'stem'|'False'`'output_code'|'True'`'sparse_output'|'False'`'max_lda_iter'|''}"

Public Const CURRENT_SETTINGS = "'Add-in'!D9"
Public Const CURRENT_TEXT_SETTINGS = "'Add-in'!D14"

Public Const BLANK_RANGE_SELECTOR = "(Click to Select)"

Public Const STATUS_CELL = "F7"
Public Const STATUS_CELL_TEXT = "F12"

Public Const WHITE = -2147483643
Public Const RED = 12632319

Public Const GRAPH_LINE_PER_INCH = 3

Public Sub prepare_for_prod()
    ThisWorkbook.Sheets("Add-in").Range("D9").Value = "{'model'|'Random forest'`'formula'|'median_property_value ~ crime_per_capita + prop_zoned_over_25k + prop_non_retail_acres + bounds_river + nox_concentration + av_rooms_per_dwelling + prop_owner_occupied_pre_1940 + dist_to_employment_ctr + highway_accessibility + tax_rate + pupil_teacher_ratio'`'param1'|'3 & 4 & 5 & 6'`'param2'|'25'`'param3'|''`'training_data'|'[xlkitlearn.xlsm]boston_housing!$A$1:$L$507'`'K'|'5'`'ts_data'|'False'`'evaluation_perc'|'30'`'evaluation_data'|''`'prediction_data'|''`'seed'|'123'`'output_model'|'True'`'output_evaluation_details'|'True'`'output_code'|'True'}"
    ThisWorkbook.Sheets("Add-in").Range("D14").Value = ""
    ThisWorkbook.Sheets("Add-in").Range("F17").Value = ""
    ThisWorkbook.Sheets("code_text").Range("D1").Value = ""
    
    Dim delete_sheets As New Collection
    Dim i As Integer
    For i = 1 To ThisWorkbook.Sheets.Count
        If (ThisWorkbook.Sheets(i).Name <> "Add-in") _
                And (ThisWorkbook.Sheets(i).Name <> "xlwings.conf") _
                And (ThisWorkbook.Sheets(i).Name <> "code_text") _
                And (ThisWorkbook.Sheets(i).Name <> "boston_housing") Then
            delete_sheets.Add ThisWorkbook.Sheets(i).Name
        End If
    Next i
    Application.DisplayAlerts = False
    For i = 1 To delete_sheets.Count
        ThisWorkbook.Sheets(delete_sheets.Item(i)).Delete
    Next i
    Application.DisplayAlerts = True
    
    ThisWorkbook.Sheets("code_text").Visible = 2
    ThisWorkbook.Sheets("xlwings.conf").Visible = 2
    
    Sheets("Add-in").CheckBoxes("chk_server").Value = -4146
    Sheets("Add-in").CheckBoxes("chk_foreground").Value = -4146
End Sub

Public Sub update_server_version()
    windows_curl "http://guetta.org/addin/update_version.php?auth_key=" & AUTH_KEY & "&new_version" & ThisWorkbook.Names("version").Value
End Sub

Public Sub update_names()
    On Error Resume Next
    
    ' Delete all names
    Dim i As Integer
    i = 0
    While (ThisWorkbook.Names.Count > 0) And (i < 50)
        ThisWorkbook.Names.Item(1).Delete
        i = i + 1
    Wend
    
    ' Store the version
    ThisWorkbook.Names.Add "version", Split(ThisWorkbook.Sheets("code_text").Range("A7"), " ")(3), False
    
    ' Delete calc_mode name if it's there
    On Error Resume Next
    ThisWorkbook.Names.Item("calc_mode").Delete
End Sub

Public Sub update_conf(Optional resolve_path As Boolean = False)
    On Error GoTo update_conf_error
    
    If SheetExists("xlwings.conf") = False Then
        Dim new_sheet_name As String
        new_sheet_name = Sheets.Add()
        Sheets(new_sheet_name).Visible = xlSheetVeryHidden
        Sheets(new_sheet_name).Name = "xlwings.conf"
    End If
    
    With Sheets("xlwings.conf")
        .Cells.Clear
        
        .Range("A1") = "Interpreter_Win"
        .Range("B1") = "%LOCALAPPDATA%\XLKitLearn\python.exe"
        
        .Range("A2") = "Interpreter_Mac"
        .Range("B2") = "$HOME/xlkitlearn/bin/python"
    End With
    
    If resolve_path = True Then
        resolve_workbook_path
    End If
    
    Exit Sub
update_conf_error:
    MsgBox "It looks like XLKitLearn has hit a snag that sometimes happens the first time you enable macros. Please save the file, " _
                & "close it, and re-open it, and you should be good to go.", vbExclamation
End Sub

