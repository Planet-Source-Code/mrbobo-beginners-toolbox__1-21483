Attribute VB_Name = "ModGP"
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    Hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (Prop As SHELLEXECUTEINFO) As Long

Public Sub GetPropDlg(frm As Form, mfile As String)
'Used to bring up the properties dialog
Dim Prop As SHELLEXECUTEINFO
Dim R As Long
With Prop
    .cbSize = Len(Prop)
    .fMask = &HC
    .Hwnd = frm.Hwnd
    .lpVerb = "properties"
    .lpFile = mfile
End With
R = ShellExecuteEx(Prop)
End Sub

