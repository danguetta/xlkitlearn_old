
Public Function dict_utils(ByVal dict_name As String, Optional ByVal Key As String = "XxXxXxXxXxXxX", Optional ByVal Value As String = "XxXxXxXxXxXxX", Optional ByVal reverse As Boolean = False, Optional ByVal error_on_missing_key As Boolean = True)
    ' This function takes a dictionary in one of the hidden names in
    ' this spreadsheet and either modifies one of the keys or returns
    ' it.
    '
    ' It takes four arguments
    '   - The name of the dictionary
    '   - The key we want to access. If this is not provided, we
    '     return an array with the keys in this dictionary
    '   - The value we want to replace it with. If this is not
    '     provided, the dictionary is left as-is, and the value
    '     corresponding to this key is returned
    '   - If reverse is included and true, the values will be used
    '     as keys, and vice versa
    '   - If error_on_missing_keys is included and False, attempts
    '     to retrieve a missing key will return the MISSING sentinel
    '     rather than throw an error (or just add the key to an
    '     existing dict)
    '
    ' Instead of the name of the dictionary, it is also possible to
    ' pass a literal dictionary to this function, in which case a
    ' value can be looked up in said dictionary, but no new value
    ' can be passed
    
    On Error GoTo error_out
    
    Dim cur_dict As String
    If InStr(dict_name, "{") <> 0 Then
        ' An actual dictionary was passed. Ensure no value was passed
        cur_dict = dict_name
    ElseIf check_valid_range(dict_name) = True Then
        cur_dict = Range(dict_name).Value
    Else
        ' We have a variable name - load it from the variable and
        ' remove the first equal sign
        cur_dict = Mid(ThisWorkbook.Names.Item(dict_name).Value, 2)
        
        ' Remove the quotes if they are there
        If Left(cur_dict, 1) = "'" Then
            cur_dict = trim_ends(cur_dict)
        End If
    End If
    
    ' Remove opening and closing braces
    cur_dict = trim_ends(cur_dict)
    
    ' We might need to construct a dictionary or list, either as the new
    ' dictionary to be output, or as a list of keys to return. Prepare
    ' a string to hold that dict/list
    Dim output_string As String
    
    ' If both the key and value are provided, output_string will need to
    ' contain the new dictionary, so we should initialize it with a
    ' brace
    If Key <> MISSING And Value <> MISSING Then
        output_string = "{"
    End If
        
    ' Go through the dictionary
    Dim split_dict, i As Integer, this_key As String, this_value As String, found_key As Boolean
    
    split_dict = Split(cur_dict, "`")
    
    For i = 0 To UBound(split_dict)
        this_key = trim_ends(Split(split_dict(i), "|")(0))
        this_value = trim_ends(Split(split_dict(i), "|")(1))
        
        If reverse = True Then
            Dim buffer As String
            buffer = this_key
            this_key = this_value
            this_value = buffer
        End If
        
        If Key = MISSING And Value = MISSING Then
            ' We're just cataloguing the keys
            output_string = output_string & this_key & "`"
        Else
            ' We're either looking to retrieve a key or set one
            If this_key = Key Then
                ' We're in the correct key - figure out what to do
                If Value = MISSING Then
                    ' We've found the value we wanted - return
                    dict_utils = this_value
                    Exit Function
                Else
                    ' Add it to the dictionary
                    output_string = output_string & "'" & Key & "'|'" & Value & "'`"
                End If
                
                ' Note that we found our key
                found_key = True
            Else
                ' If we're just looking for a key, don't do anything. Otherwise, just
                ' leave this key as is in the output dictionary
                If Value <> MISSING Then
                    output_string = output_string & "'" & this_key & "'|'" & this_value & "'`"
                End If
            End If
        End If
    Next i
    
    ' If we got here, it's for one of the following reasons
    '   - We're retrieving a key and it wasn't found
    '   - We're updating a dictionary
    '   - We want to return a list of keys
    
    ' First, check if the key wasn't found
    If Key <> MISSING And Value = MISSING Then
        ' Retrieving a key that wasn't found
        If error_on_missing_key Then
            MsgBox "CODING ERROR: attempting to retrieve non-existent key " & Key, vbCritical
            GoTo error_out
        Else
            dict_utils = MISSING
        End If
        
        Exit Function
    ElseIf Key <> MISSING And Value <> MISSING Then
        ' Updating a dictionary. Make sure we've found our key, and
        ' add it if needed
        If found_key = False Then
            If error_on_missing_key Then
                MsgBox "CODING ERROR: attempting to set non-existent key " & Key, vbCritical
                GoTo error_out
            Else
                ' Add the key
                output_string = output_string & "'" & Key & "'|'" & Value & "'`"
            End If
        End If
    End If
    
    ' There will be an extra backtick at the end; remove it
    output_string = Mid(output_string, 1, Len(output_string) - 1)
    
    ' Now, find out what we're trying to do
    If Key <> MISSING And Value <> MISSING Then
        ' Add a closing brace
        output_string = output_string & "}"
    
        ' Save it to the workbook or return it
        If InStr(dict_name, "{") <> 0 Then
            dict_utils = output_string
        ElseIf check_valid_range(dict_name) = True Then
            Range(dict_name).Value = output_string
        Else
            ThisWorkbook.Names.Item(dict_name).Value = output_string
        End If
    Else
        ' We need to output our keys
        dict_utils = Split(output_string, "`")
        Exit Function
    End If
    
    Exit Function
error_out:
    dict_utils = Split("", ",")
End Function

Public Function sync_dicts(ByVal new_dict As String, ByVal template_dict As String)
    Dim Keys
    Dim i As Integer
    Dim output_dict As String
    Dim this_val As String
    
    ' Retrieve the keys from the template_dict
    Keys = dict_utils(template_dict)
    
    ' Loop through them and add each entry. Start with a blank dict
    output_dict = "{}"
    For i = 0 To UBound(Keys)
        this_val = dict_utils(new_dict, Keys(i), , , False)
        If this_val = MISSING Then
            output_dict = dict_utils(output_dict, Keys(i), dict_utils(template_dict, Keys(i)), , False)
        Else
            output_dict = dict_utils(output_dict, Keys(i), this_val, , False)
        End If
    Next i
    
    sync_dicts = output_dict
End Function