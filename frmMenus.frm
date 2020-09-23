VERSION 5.00
Begin VB.Form frmMenus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Fun"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Example usage"
      Height          =   1455
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Apply to NotePad.exe"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update NotePad.exe"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Effected Menus (Backgrounds)"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
      Begin VB.CheckBox Check1 
         Caption         =   "Include Menu Bar"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Menu Items"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Single Menu Item"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu Background Settings"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMenus.frx":0000
         Left            =   240
         List            =   "frmMenus.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMenus.frx":001C
         Left            =   1920
         List            =   "frmMenus.frx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMenus.frx":0053
         Left            =   1320
         List            =   "frmMenus.frx":0069
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Hatch style :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Default Settings"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuFatman 
         Caption         =   "Fatman"
      End
      Begin VB.Menu mnuRatboy 
         Caption         =   "Ratboy"
      End
   End
   Begin VB.Menu mnuMenu2 
      Caption         =   "Menu2"
      Begin VB.Menu mnufish 
         Caption         =   "Fish"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheColor As Long
Dim TheStyle As Long
Dim DontUpdate As Boolean
Dim sfile As String

Private Sub Check1_Click()
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
    Case 0
        TheStyle = 0
        Combo3.Enabled = False
    Case 1
        TheStyle = 2
        Combo3.Enabled = True
End Select
If DontUpdate Then Exit Sub
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
    Case 0
        TheColor = vbRed
    Case 1
        TheColor = vbGreen
    Case 2
        TheColor = vbBlue
    Case 3
        TheColor = vbYellow
    Case 4
        TheColor = vbCyan
End Select
If DontUpdate Then Exit Sub
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex


End Sub


Private Sub Combo3_Click()
If DontUpdate Then Exit Sub
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub


Private Sub Combo4_Click()
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub

Private Sub Command1_Click()
Dim NtPWnd As Long
Dim hMenu As Long, hSubMenu As Long, nCnt As Long
Shell "c:\windows\notepad.exe", vbNormalFocus
NtPWnd = FindWindow(vbNullString, "Untitled - Notepad")
SetMenuBackColor NtPWnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub


Private Sub Command2_Click()
SetDefs Me.Hwnd
End Sub

Private Sub Command3_Click()
Dim NtPWnd As Long
Dim hMenu As Long, hSubMenu As Long, nCnt As Long
NtPWnd = FindWindow(vbNullString, "Untitled - Notepad")
SetMenuBackColor NtPWnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
DontUpdate = True
For x = 0 To 50
    Combo4.AddItem Str(x)
Next x
Combo4.ListIndex = 0
Combo1.ListIndex = 0
Combo2.ListIndex = 0
DontUpdate = False
Combo3.ListIndex = 0
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuMenu
    End If
End Sub

Private Sub Option1_Click()
Combo4.Enabled = Option2.Value
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub

Private Sub Option2_Click()
Combo4.Enabled = Option2.Value
SetMenuBackColor Me.Hwnd, TheColor, TheStyle, Combo3.ListIndex, Option2.Value, Combo4.ListIndex

End Sub
