VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_data_errors 
   Caption         =   "Data errors"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915.001
   OleObjectBlob   =   "frm_data_errors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_data_errors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    lst_errors.ColumnCount = 5
    lst_errors.ColumnWidths = "50,15,170,15,500"
    
    Dim i As Integer
    For i = 1 To header_errors.count()
        lst_errors.AddItem
        
        lst_errors.List(i - 1, 0) = Split(header_errors.Item(i), "`")(0)
        lst_errors.List(i - 1, 1) = ""
        lst_errors.List(i - 1, 2) = Split(header_errors.Item(i), "`")(1)
        lst_errors.List(i - 1, 3) = ""
        lst_errors.List(i - 1, 4) = Split(header_errors.Item(i), "`")(2)
    Next i
    
End Sub
