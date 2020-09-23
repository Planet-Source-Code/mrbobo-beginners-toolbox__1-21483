VERSION 5.00
Begin VB.Form frmWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Task List"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Close"
      Height          =   375
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Hide"
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Show"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Restore"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Maximise"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Minimise"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API
'Find a window by its caption
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Send a message to the window - used to close the window
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Used by the LoadTaskList function to get a list of running tasks
Private Declare Function GetWindow Lib "user32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
'Set the desired window state for a window
Private Declare Function ShowWindow Lib "user32.dll" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Const SW_HIDE = 0
Const SW_MAXIMIZE = 3
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_RESTORE = 9

Private Sub cmdAction_Click(Index As Integer)
Dim winHwnd As Long
Dim RetVal As Long
'Find the window by its caption as determined by
'the LoadTaskList function
winHwnd = FindWindow(vbNullString, Combo1.Text)
If winHwnd <> 0 Then
    Select Case Index
        Case 0
            RetVal = ShowWindow(winHwnd, SW_MINIMIZE)
        Case 1
            RetVal = ShowWindow(winHwnd, SW_MAXIMIZE)
        Case 2
            RetVal = ShowWindow(winHwnd, SW_RESTORE)
        Case 3
            RetVal = ShowWindow(winHwnd, SW_SHOW)
        Case 4
            RetVal = ShowWindow(winHwnd, SW_HIDE)
        Case 5
            RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
            'We just closed the window so update the combo
            LoadTaskList
    End Select
End If
End Sub

Sub LoadTaskList()
'This function (not mine) returns all tasks running
'and puts its window caption in the combo box
Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim parent As Long
Combo1.Clear
CurrWnd = GetWindow(Form1.Hwnd, GW_HWNDFIRST)
While CurrWnd <> 0
parent = GetParent(CurrWnd)
Length = GetWindowTextLength(CurrWnd)
TaskName = Space$(Length + 1)
Length = GetWindowText(CurrWnd, TaskName, Length + 1)
TaskName = Left$(TaskName, Len(TaskName) - 1)
If Length > 0 Then
    If TaskName <> Me.Caption Then
        If TaskName <> Form1.Caption Then
            If TaskName <> "taskmon" Then
                Combo1.AddItem TaskName
            End If
        End If
    End If
End If
CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
DoEvents
Wend
If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub
Private Sub Command1_Click()
LoadTaskList
End Sub
Private Sub Form_Load()
LoadTaskList
End Sub
