Attribute VB_Name = "mdl_telemetry"
Option Explicit

Public Sub windows_curl(request As String)
    ' This function runs a curl on windows
    
    ' Replace spaces in the request with '%20' and double quotes with single quotes
    request = Replace(Replace(request, " ", "%20"), """", "'")
    
    ' Add the curl to the front of the request, with double quotes around the request
    request = "curl """ & request & """"
    
    ' Run without waiting for return
    
    Dim WaitOnReturn As Boolean: WaitOnReturn = False
    Dim WindowStyle As Integer: WindowStyle = 0
    Dim Wsh As Object
    Set Wsh = CreateObject("WScript.Shell")
    Wsh.Run request, WindowStyle, WaitOnReturn
    Set Wsh = Nothing
End Sub

Public Sub log_vba_error(Content As String)
    ' Log a VBA error

    On Error Resume Next
    Dim request As String
    
    #If Mac Then
        request = Replace(Content, "'", """")
        RunPython "import requests; requests.post(url = 'http://guetta.org/addin/error.php'," & _
                        "data = {'run_id':'" & run_id() & "', " & _
                        "        'source':'unknown', 'error_type'='vba_error', 'platform'='mac', " & _
                        "        'error_text' : '" & request & "'}, timeout = 10)"
    #Else
        request = Replace(Content, vbCrLf, "\n")
        request = Replace(request, Chr(10), "\n")
        request = Replace(request, "&", "|")
        
        windows_curl "http://guetta.org/addin/error.php?run_id=" & run_id() & _
                           "&source=unknown&error_type=vba_error&platform=windows&error_text=" & request
    #End If

End Sub
