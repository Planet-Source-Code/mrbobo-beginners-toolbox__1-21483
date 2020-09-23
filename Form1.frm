VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Starter Kit Volume 1"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   4290
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3413
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   2e6
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Strings"
      TabPicture(0)   =   "Form1.frx":00E3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "Form1.frx":00FF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Navigation"
      TabPicture(2)   =   "Form1.frx":011B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Read/Write"
      TabPicture(3)   =   "Form1.frx":0137
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Examples"
      TabPicture(4)   =   "Form1.frx":0153
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame8 
         Height          =   3495
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   6615
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "     C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe     "
            Top             =   360
            Width           =   4695
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Instr"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   61
            Top             =   1500
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "InstrRev"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   60
            Top             =   1980
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Trim"
            Height          =   375
            Index           =   2
            Left            =   360
            TabIndex        =   59
            Top             =   2460
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Split"
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   58
            Top             =   2940
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Left"
            Height          =   375
            Index           =   4
            Left            =   2400
            TabIndex        =   57
            Top             =   1500
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Mid"
            Height          =   375
            Index           =   5
            Left            =   2400
            TabIndex        =   56
            Top             =   1980
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Right"
            Height          =   375
            Index           =   6
            Left            =   2400
            TabIndex        =   55
            Top             =   2460
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Join"
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   54
            Top             =   2940
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Path Only"
            Height          =   375
            Index           =   8
            Left            =   4440
            TabIndex        =   53
            Top             =   1500
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "File Only"
            Height          =   375
            Index           =   9
            Left            =   4440
            TabIndex        =   52
            Top             =   1980
            Width           =   1815
         End
         Begin VB.CommandButton cmdStrings 
            Caption         =   "Extension Only"
            Height          =   375
            Index           =   10
            Left            =   4440
            TabIndex        =   51
            Top             =   2460
            Width           =   1815
         End
         Begin VB.Label Label1 
            Height          =   615
            Left            =   360
            TabIndex        =   63
            Top             =   780
            Width           =   5895
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   43
         Top             =   600
         Width           =   6495
         Begin VB.CommandButton Command12 
            Caption         =   "Window handler"
            Height          =   375
            Left            =   480
            TabIndex        =   64
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "File Properties"
            Height          =   375
            Left            =   480
            TabIndex        =   46
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Text Editor"
            Height          =   375
            Left            =   480
            TabIndex        =   45
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Menus"
            Height          =   375
            Left            =   480
            TabIndex        =   44
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Change the state of any open windows"
            Height          =   255
            Left            =   2400
            TabIndex        =   65
            Top             =   2520
            Width           =   3375
         End
         Begin VB.Label Label6 
            Caption         =   "Brings up the standard Windows File Property Dialog"
            Height          =   375
            Left            =   2400
            TabIndex        =   49
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label Label5 
            Caption         =   "A basic text editor showing toolbar functions, cut, copy, paste, infinite undo/redo, using temp files."
            Height          =   495
            Left            =   2400
            TabIndex        =   48
            Top             =   1320
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   "Change the appearance of the standard VB menu"
            Height          =   375
            Left            =   2400
            TabIndex        =   47
            Top             =   720
            Width           =   3855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Images"
         Height          =   3495
         Left            =   -71400
         TabIndex        =   31
         Top             =   480
         Width           =   3255
         Begin VB.HScrollBar HS 
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2640
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.VScrollBar VS 
            Height          =   2295
            Left            =   2760
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox PicFrame 
            Height          =   2535
            Left            =   240
            ScaleHeight     =   2475
            ScaleWidth      =   2715
            TabIndex        =   39
            Top             =   360
            Width           =   2775
            Begin VB.PictureBox Picture1 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   0
               ScaleHeight     =   735
               ScaleWidth      =   1095
               TabIndex        =   40
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Save As"
            Height          =   375
            Left            =   1680
            TabIndex        =   35
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Open"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   3000
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Text"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   30
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton Option4 
            Caption         =   "Rich Text Format"
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   2640
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Plain Text"
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   2400
            Value           =   -1  'True
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   1935
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3413
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Form1.frx":016F
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Save As"
            Height          =   375
            Left            =   1680
            TabIndex        =   33
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Open"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   3000
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   23
         Top             =   480
         Width           =   6375
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":0340
            Left            =   360
            List            =   "Form1.frx":0350
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   840
            Width           =   2535
         End
         Begin VB.FileListBox File1 
            Height          =   2820
            Left            =   3360
            TabIndex        =   26
            Top             =   360
            Width           =   2655
         End
         Begin VB.DirListBox Dir1 
            Height          =   1890
            Left            =   360
            TabIndex        =   25
            Top             =   1320
            Width           =   2535
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Browse for folder"
         Height          =   975
         Left            =   -70680
         TabIndex        =   20
         Top             =   3000
         Width           =   2415
         Begin VB.CommandButton Command2 
            Caption         =   "Browse"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Without MS Shell"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Common Dialog"
         Height          =   2295
         Left            =   -70680
         TabIndex        =   13
         Top             =   600
         Width           =   2415
         Begin VB.CommandButton cmdCmnDlg 
            Caption         =   "Font"
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   19
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdCmnDlg 
            Caption         =   "Color"
            Height          =   375
            Index           =   2
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdCmnDlg 
            Caption         =   "Save"
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   17
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdCmnDlg 
            Caption         =   "Open"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   855
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   2040
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Pure code"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "With MS Common Dialog Control 6.0"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Using MS Shell Controls and Automation"
         Height          =   3375
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   3855
         Begin VB.CommandButton cmd 
            Caption         =   "Browse for folder"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3375
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Browse for folder (Root : Desktop)"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   3375
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Open Windows folder"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   3375
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Find Files"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Run File"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   8
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Tray Properties"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   7
            Top             =   2880
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Minimise All"
            Height          =   375
            Index           =   6
            Left            =   2040
            TabIndex        =   6
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Un Minimise All"
            Height          =   375
            Index           =   7
            Left            =   2040
            TabIndex        =   5
            Top             =   2400
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Code :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I have not commented this code because the useful
'code is shown in the RichtextBox
Dim MyShell As shell32.Shell
Dim sfile As String

Private Sub cmd_Click(Index As Integer)
On Error Resume Next
sfile = ""
Select Case Index
    Case 0
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   On Error Resume Next" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   sfile = MyShell.BrowseForFolder(Me.hWnd, " + Chr(34) + "fish" + Chr(34) + ", 0)" + vbCrLf + _
        "End Sub"
        sfile = MyShell.BrowseForFolder(Me.Hwnd, "fish", 0)
        If sfile <> "" Then
            MsgBox "You selected " + sfile
        End If

    Case 1
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   On Error Resume Next" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   sfile = MyShell.BrowseForFolder(Me.hWnd, " + Chr(34) + "fish" + Chr(34) + ", 0," + Chr(34) + "C:\windows\desktop" + Chr(34) + ")" + vbCrLf + _
        "End Sub"
        sfile = MyShell.BrowseForFolder(Me.Hwnd, "fish", 0, "C:\windows\desktop")
        If sfile <> "" Then
            MsgBox "You selected " + sfile
        End If
    Case 2
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.Open (" + Chr(34) + "C:\windows" + Chr(34) + ")" + vbCrLf + _
        "End Sub"
        MyShell.Open ("C:\windows")
    Case 3
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.FindFiles" + vbCrLf + _
        "End Sub"
        MyShell.FindFiles
    Case 4
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.FileRun" + vbCrLf + _
        "End Sub"
        MyShell.FileRun
    Case 5
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.TrayProperties" + vbCrLf + _
        "End Sub"
        MyShell.TrayProperties
    Case 6
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.MinimizeAll" + vbCrLf + _
        "End Sub"
        MyShell.MinimizeAll
    Case 7
        RTF.Text = "'You need to reference 'MS Shell Controls and Automation' in this project" + vbCrLf + _
        "Dim MyShell As Shell32.Shell" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "   Dim sfile As String" + vbCrLf + _
        "   Set MyShell = New Shell32.Shell" + vbCrLf + _
        "   MyShell.UndoMinimizeALL" + vbCrLf + _
        "End Sub"
        MyShell.UndoMinimizeALL
End Select
End Sub

Private Sub cmdCmnDlg_Click(Index As Integer)
On Error GoTo woops
cmdCmnDlg(2).BackColor = &H8000000F
With cmdCmnDlg(3)
    .FontName = "MS Sans Serif"
    .FontSize = 8
    .FontBold = False
    .FontItalic = False
    .FontStrikethru = False
    .FontUnderline = False
End With
If Option1.Value = True Then
    Select Case Index
        Case 0
            RTF.Text = "Dim sfile As String" + vbCrLf + _
            "On Error GoTo woops" + vbCrLf + _
            "With CommonDialog1" + vbCrLf + _
            "    .DialogTitle = " + Chr(34) + "Open Files" + Chr(34) + vbCrLf + _
            "    .CancelError = True" + vbCrLf + _
            "    .Filter = " + Chr(34) + "Text files (*.txt;*.doc)|*.txt;*.doc|All files (*.*)|*.*" + Chr(34) + vbCrLf + _
            "    .ShowOpen" + vbCrLf + _
            "     If Len(.FileName) = 0 Then Exit Sub" + vbCrLf + _
            "    sfile = .FileName" + vbCrLf + _
            "End With" + vbCrLf + _
            "woops: Exit Sub"
            With CommonDialog1
                .DialogTitle = "Open Files"
                .CancelError = True
                .Filter = "Text files (*.txt;*.doc)|*.txt;*.doc|All files (*.*)|*.*"
                .ShowOpen
                If Len(.FileName) = 0 Then Exit Sub
                sfile = .FileName
            End With
            MsgBox "You selected " + sfile + vbCrLf + "You need to add code to actually open the file."
        Case 1
            RTF.Text = "Dim sfile As String" + vbCrLf + _
            "On Error GoTo woops" + vbCrLf + _
            "With CommonDialog1" + vbCrLf + _
            "    .DialogTitle = " + Chr(34) + "Save Files" + Chr(34) + vbCrLf + _
            "    .CancelError = True" + vbCrLf + _
            "    .Filter = " + Chr(34) + "Text files (*.txt;*.doc)|*.txt;*.doc|All files (*.*)|*.*" + Chr(34) + vbCrLf + _
            "    .ShowSave" + vbCrLf + _
            "     If Len(.FileName) = 0 Then Exit Sub" + vbCrLf + _
            "    sfile = .FileName" + vbCrLf + _
            "End With" + vbCrLf + _
            "woops: Exit Sub"
            With CommonDialog1
                .DialogTitle = "Save Files"
                .CancelError = True
                .Filter = "Text files (*.txt;*.doc)|*.txt;*.doc|All files (*.*)|*.*"
                .ShowSave
                If Len(.FileName) = 0 Then Exit Sub
                sfile = .FileName
            End With
            MsgBox "You selected " + sfile + vbCrLf + "You need to add code to actually save the file."
        Case 2
            RTF.Text = "CommonDialog1.CancelError = True" + vbCrLf + _
            "CommonDialog1.flags = 0" + vbCrLf + _
            "CommonDialog1.Action = 3" + vbCrLf + _
            "Me.BackColor = CommonDialog1.Color"
            CommonDialog1.CancelError = True
            CommonDialog1.flags = 0
            CommonDialog1.Action = 3
            cmdCmnDlg(2).BackColor = CommonDialog1.Color
        Case 3
            RTF.Text = "CommonDialog1.flags = cdlCFBoth Or cdlCFEffects" + vbCrLf + _
            "CommonDialog1.ShowFont" + vbCrLf + _
            "With RichTextBox1" + vbCrLf + _
            "    .SelFontName = CommonDialog1.FontName" + vbCrLf + _
            "    .SelFontSize = CommonDialog1.FontSize" + vbCrLf + _
            "    .SelBold = CommonDialog1.FontBold" + vbCrLf + _
            "    .SelItalic = CommonDialog1.FontItalic" + vbCrLf + _
            "    .SelStrikeThru = CommonDialog1.FontStrikethru" + vbCrLf + _
            "    .SelUnderline = CommonDialog1.FontUnderline" + vbCrLf + _
            "    .SelColor = CommonDialog1.Color" + vbCrLf + _
            "End With"
            CommonDialog1.flags = cdlCFBoth Or cdlCFEffects
            CommonDialog1.ShowFont
            With cmdCmnDlg(3)
                .FontName = CommonDialog1.FontName
                .FontSize = CommonDialog1.FontSize
                .FontBold = CommonDialog1.FontBold
                .FontItalic = CommonDialog1.FontItalic
                .FontStrikethru = CommonDialog1.FontStrikethru
                .FontUnderline = CommonDialog1.FontUnderline
            End With
    End Select
ElseIf Option2.Value = True Then
    Select Case Index
        Case 0
            RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
            "sfile = ShowOpen(" + Chr(34) + "Text Files (*.txt;*.doc)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt;*.doc" + Chr(34) + " + Chr(0) + " + Chr(34) + "All Files (*.*)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.*" + Chr(34) + " + Chr(0), 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")"
            sfile = ShowOpen("Text Files (*.txt;*.doc)" + Chr(0) + "*.txt;*.doc" + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
        Case 1
            RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
            "sfile = ShowSave(" + Chr(34) + "Text Files (*.txt;*.doc)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt;*.doc" + Chr(34) + " + Chr(0) + " + Chr(34) + "All Files (*.*)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.*" + Chr(34) + " + Chr(0), 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")"
            sfile = ShowSave("Text Files (*.txt;*.doc)" + Chr(0) + "*.txt;*.doc" + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
        Case 2
            RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
            "Private Sub Form_Load()" + vbCrLf + _
            "   InitCmnDlg Me.hwnd" + vbCrLf + _
            "   Me.BackColor = ShowColor" + vbCrLf + _
            "End Sub"
            cmdCmnDlg(2).BackColor = ShowColor
        Case 3
            RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
            "ShowFont" + vbCrLf + _
            "With RichTextBox1" + vbCrLf + _
            "   .SelFontName = mFontName" + vbCrLf + _
            "   .SelFontSize = mFontsize" + vbCrLf + _
            "   .SelBold = mBold" + vbCrLf + _
            "   .SelItalic = mItalic" + vbCrLf + _
            "   .SelStrikeThru = mStrikethru" + vbCrLf + _
            "   .SelUnderline = mUnderline" + vbCrLf + _
            "   .SelColor = mFontColor" + vbCrLf + _
            "End With"
            ShowFont
            With cmdCmnDlg(3)
                .FontName = mFontName
                .FontSize = mFontsize
                .FontBold = mBold
                .FontItalic = mItalic
                .FontStrikethru = mStrikethru
                .FontUnderline = mUnderline
            End With

    End Select
End If
woops:
End Sub

Private Sub cmdStrings_Click(Index As Integer)
Dim found As Long, temp As Variant
Select Case Index
    Case 0
       found = InStr(1, Text1.Text, "VB")
       If found <> 0 Then
            Label1 = "Found VB at character " + Str(found)
            Text1.SetFocus
            Text1.SelStart = found - 1
            Text1.SelLength = 2
       End If
       RTF.Text = "found = InStr(1, Text1.Text, " + Chr(34) + "VB" + Chr(34) + ")" + vbCrLf + _
                    "If found <> 0 Then" + vbCrLf + _
                    "     Text1.SetFocus" + vbCrLf + _
                    "     Text1.SelStart = found - 1" + vbCrLf + _
                    "     Text1.SelLength = 2" + vbCrLf + _
                    "End If"
    Case 1
       found = InStrRev(Text1.Text, "VB")
       If found <> 0 Then
            Label1 = "Found VB at character " + Str(found)
            Text1.SetFocus
            Text1.SelStart = found - 1
            Text1.SelLength = 2
       End If
       RTF.Text = "found = InStrRev(Text1.Text, " + Chr(34) + "VB" + Chr(34) + ")" + vbCrLf + _
                    "If found <> 0 Then" + vbCrLf + _
                    "     Text1.SetFocus" + vbCrLf + _
                    "     Text1.SelStart = found - 1" + vbCrLf + _
                    "     Text1.SelLength = 2" + vbCrLf + _
                    "End If"
    Case 2
        Text1.Text = Trim(Text1.Text)
        Label1 = "The Trim function removes spaces from the start and end of a string."
        RTF.Text = "Text1.Text = Trim(Text1.Text)"
    Case 3
        Label1 = ""
        For x = 0 To UBound(Split(Text1.Text, "\"))
            Label1 = Label1 + Split(Text1.Text, "\")(x) + ", "
        Next x
        Label1 = Trim(Label1)
        If Right(Label1, 1) = "," Then Label1 = Left(Label1, Len(Label1) - 1)
        Label1 = "Divided this string into" + Str(x) + " compoents - " + Label1
        RTF.Text = "Dim fred() as string" + vbCrLf + _
        "For x = 0 To UBound(Split(Text1.Text, " + Chr(34) + "\" + Chr(34) + "))" + vbCrLf + _
        "   fred(x) = Split(Text1.Text, " + Chr(34) + "\" + Chr(34) + ")" + vbCrLf + _
        "Next x"
    Case 4
        Label1 = "The 10 characters on the left of this string are - " + Left(Trim(Text1.Text), 10)
        RTF.Text = "Text1.Text = Left(Text1.Text, 10)"
    Case 5
        Label1 = "The characters in the middle of this string, starting at the tenth letter and ending at the tenth last letter are - " + Mid(Trim(Text1.Text), 11, Len(Trim(Text1.Text)) - 20)
        RTF.Text = "Text1.Text = Mid(Text1.Text, 11, Len(Text1.Text) - 20)"
    Case 6
        Label1 = "The 10 characters on the right of this string are - " + Right(Trim(Text1.Text), 10)
        RTF.Text = "Text1.Text = Right(Text1.Text, 10)"
    Case 7
        temp = Split(Text1.Text, "Microsoft Visual Studio")
        Label1 = "The Join function complements the Split function. It combines an array of strings with a chosen delimiter." + vbCrLf + Join(temp, "Bobo Enterprises")
        RTF.Text = "Dim temp As Variant" + vbCrLf + _
        "temp = Split(Text1.Text, " + Chr(34) + "Microsoft Visual Studio" + Chr(34) + ")" + vbCrLf + _
        "Label1 = Join(temp, " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")"
    Case 8
        Label1 = Mid(Trim(Text1.Text), 1, InStrRev(Trim(Text1.Text), "\"))
        RTF.Text = "Label1 = Mid(Text1.Text, 1, InStrRev(Trim(Text1.Text), " + Chr(34) + "\" + Chr(34) + "))"
    Case 9
        Label1 = Mid(Trim(Text1.Text), InStrRev(Trim(Text1.Text), "\") + 1)
        RTF.Text = "Label1 = Mid(Text1.Text, InStrRev(Text1.Text, " + Chr(34) + "\" + Chr(34) + ") + 1)"
    Case 10
        Label1 = Mid(Trim(Text1.Text), InStrRev(Trim(Text1.Text), ".") + 1)
        RTF.Text = "Label1 = Mid(Text1.Text, InStrRev(Text1.Text, " + Chr(34) + "." + Chr(34) + ") + 1)"
End Select
End Sub



Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText RTF.Text
End Sub

Private Sub Command10_Click()
frmTxtEdit.Show vbModal
End Sub

Private Sub Command11_Click()
sfile = ShowOpen("All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
If sfile <> "" Then GetPropDlg Me, sfile


End Sub

Private Sub Command12_Click()
frmWindow.Show vbModal
End Sub

Private Sub Command2_Click()
RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project." + vbCrLf + _
"sfile = BrowseForFolder(" + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")"
sfile = BrowseForFolder("C:\windows", "Bobo Enterprises")
If sfile <> "" Then MsgBox "You selected " + sfile
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form2.Show vbModal
End Sub

Private Sub Command5_Click()
If Option3.Value = True Then
    RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
    "sfile = ShowOpen(" + Chr(34) + "Text Files (*.txt)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt" + Chr(34) + " + Chr(0) + " + Chr(34) + "All Files (*.*)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.*" + Chr(34) + " + Chr(0), 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
    "If sfile <> " + Chr(34) + Chr(34) + " Then RichTextBox1.Text = ReadText(sfile)"
    sfile = ShowOpen("Text Files (*.txt)" + Chr(0) + "*.txt" + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
    If sfile <> "" Then RichTextBox1.Text = ReadText(sfile)
ElseIf Option4.Value = True Then
    RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
    "sfile = ShowOpen(" + Chr(34) + "Text Files (*.txt)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt" + Chr(34) + " + Chr(0) + " + Chr(34) + "All Files (*.*)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.*" + Chr(34) + " + Chr(0), 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
    "If sfile <> " + Chr(34) + Chr(34) + " Then RichTextBox1.LoadFile sfile"
    sfile = ShowOpen("Text Files (*.txt)" + Chr(0) + "*.txt" + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
    If sfile <> "" Then RichTextBox1.LoadFile sfile
End If
End Sub

Private Sub Command6_Click()
If Option3.Value = True Then
    RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
    "sfile = ShowSave(" + Chr(34) + "Text Files (*.txt)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt" + Chr(34) + " + Chr(0) + " + Chr(34) + ", 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
    "If sfile <> " + Chr(34) + Chr(34) + " Then FileSave RichTextBox1.Text, sfile"
    sfile = ShowSave("Text Files (*.txt)" + Chr(0) + "*.txt" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
    If sfile <> "" Then
        If Right(sfile, 4) <> ".txt" Then sfile = sfile + ".txt"
        FileSave RichTextBox1.Text, sfile
    End If
ElseIf Option4.Value = True Then
    RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
    "sfile = ShowSave(" + Chr(34) + "Text Files (*.txt)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.txt" + Chr(34) + " + Chr(0) + " + Chr(34) + ", 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
    "If sfile <> " + Chr(34) + Chr(34) + " Then RichTextBox1.SaveFile sfile"
    sfile = ShowSave("Text Files (*.txt)" + Chr(0) + "*.txt" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
    If sfile <> "" Then
        If Right(sfile, 4) <> ".txt" Then sfile = sfile + ".txt"
        RichTextBox1.SaveFile sfile
    End If
End If

End Sub

Private Sub Command7_Click()
RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
"sfile = ShowOpen(" + Chr(34) + "Image Files (*.bmp;*.jpg;*.gif)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.bmp;*.jpg;*.gif" + Chr(34) + " + Chr(0) + " + Chr(34) + "All Files (*.*)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.*" + Chr(34) + " + Chr(0), 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
"If sfile <> " + Chr(34) + Chr(34) + " Then Picture1.Picture = LoadPicture(sfile)"
sfile = ShowOpen("Image Files (*.bmp;*.jpg;*.gif)" + Chr(0) + "*.bmp;*.jpg;*.gif" + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
If sfile <> "" Then Picture1.Picture = LoadPicture(sfile)
SizeImage
End Sub

Private Sub Command8_Click()
RTF.Text = "'You need to include the Module 'CmnDialog.bas' in your project or use MS CommonDialog Control" + vbCrLf + _
"sfile = ShowSave(" + Chr(34) + "Bitmap Files (*.bmp)" + Chr(34) + " + Chr(0) + " + Chr(34) + "*.bmp" + Chr(34) + " + Chr(0) + " + Chr(34) + ", 5, " + Chr(34) + "C:\windows" + Chr(34) + ", " + Chr(34) + "Bobo Enterprises" + Chr(34) + ")" + vbCrLf + _
"If sfile <> " + Chr(34) + Chr(34) + " Then SavePicture Picture1.Image, sfile"
sfile = ShowSave("Bitmap Files (*.bmp)" + Chr(0) + "*.bmp" + Chr(0), 5, "C:\windows", "Bobo Enterprises")
If sfile <> "" Then
    If Right(sfile, 4) <> ".bmp" Then sfile = sfile + ".bmp"
    SavePicture Picture1.Image, sfile
End If

End Sub

Private Sub Command9_Click()
frmMenus.Show vbModal
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.path = Drive1.Drive
File1.path = Dir1.path
End Sub

Private Sub Form_Load()
Set MyShell = New shell32.Shell
InitCmnDlg Me.Hwnd
Combo1.ListIndex = 3
End Sub
Private Sub Combo1_Click()
On Error Resume Next

Select Case Combo1.ListIndex
Case 0
    File1.Pattern = "*.DOC"
Case 1
    File1.Pattern = "*.TXT"
Case 2
    File1.Pattern = "*.TXT;*.DOC;*.WRI;*.INI;*.INF;*.LOG;*.RTF"
Case 3
    File1.Pattern = "*.*"
End Select

End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Unload frmMenus
Unload frmTxtEdit
Unload frmWindow

End Sub

Private Sub HS_Change()
Picture1.Left = -HS.Value
End Sub

Private Sub HS_Scroll()
Picture1.Left = -HS.Value

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
    RTF.Text = "Private Sub Dir1_Change()" + vbCrLf + "    File1.Path = Dir1.Path" + vbCrLf + "End Sub" + vbCrLf + vbCrLf + _
    "Private Sub Drive1_Change()" + vbCrLf + "On Error Resume Next" + vbCrLf + "Dir1.Path = Drive1.Drive" + vbCrLf + "File1.Path = Dir1.Path" + vbCrLf + "End Sub" + vbCrLf + vbCrLf + _
    "Private Sub Form_Load()" + vbCrLf + "Set MyShell = New shell32.Shell" + vbCrLf + "InitCmnDlg Me.hwnd" + vbCrLf + "Combo1.ListIndex = 3" + vbCrLf + _
    "End Sub" + vbCrLf + vbCrLf + "Private Sub Combo1_Click()" + vbCrLf + "On Error Resume Next" + vbCrLf + "Select Case Combo1.ListIndex" + vbCrLf + "  Case 0" + vbCrLf + "      File1.Pattern = " + Chr(34) + "*.DOC" + Chr(34) + vbCrLf + _
    "  Case 1" + vbCrLf + "      File1.Pattern = " + Chr(34) + "*.TXT" + Chr(34) + vbCrLf + "  Case 2" + vbCrLf + "      File1.Pattern = " + Chr(34) + "*.TXT;*.DOC;*.WRI;*.INI;*.INF;*.LOG;*.RTF" + Chr(34) + vbCrLf + _
    "  Case 3" + vbCrLf + "      File1.Pattern = " + Chr(34) + "*.*" + Chr(34) + vbCrLf + "  End Select" + vbCrLf + "End Sub"
End If
End Sub

Public Sub FileSave(Text As String, filepath As String)
On Error Resume Next
Dim Directory As String
              Directory$ = filepath
              Open Directory$ For Output As #1
              
           Print #1, Text
       Close #1
Exit Sub
End Sub
Public Function ReadText(path As String) As String
    Dim temptxt As String
    temptxt = ""
    Open path For Input As #1
    Do While Not EOF(1)
        Line Input #1, x
        temptxt = temptxt + x + vbCrLf
    Loop
    Close #1
    ReadText = temptxt
End Function


Public Sub SizeImage()
If Picture1.Width > PicFrame.Width Then
    HS.Visible = True
    PicFrame.Height = 2535 - HS.Height
Else
    PicFrame.Height = 2535
    HS.Visible = False
End If
If Picture1.Height > PicFrame.Height Then
    VS.Visible = True
    PicFrame.Width = PicFrame.Width - VS.Width
Else
    PicFrame.Width = 2775
    VS.Visible = False
End If
HS.Width = PicFrame.Width
VS.Height = PicFrame.Height
HS.Max = Picture1.Width - HS.Width
VS.Max = Picture1.Height - VS.Height
End Sub

Private Sub VS_Change()
Picture1.Top = -VS.Value

End Sub

Private Sub VS_Scroll()
Picture1.Top = -VS.Value

End Sub
