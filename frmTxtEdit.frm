VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTxtEdit 
   Caption         =   "Bobo Basic Text Editor"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7440
   Icon            =   "frmTxtEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgTB 
      Left            =   6240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0556
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0672
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":078E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":08AA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":09C6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0AE2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0BFE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTxtEdit.frx":0D12
            Key             =   "Redo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmTxtEdit.frx":0E26
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileMRUSpace 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmTxtEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTempFilename Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFilename As String _
    ) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal Hwnd As Long) As Long
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Dim curfile As String 'Full path of file in the rich textbox
Dim curEXT As String 'Example "txt"
Dim lastopened As String 'Path to last opened from dir
Dim lastopenedEXT As String 'Extension of last opened file
Dim lastsaved As String 'Path to last saved to dir
Dim MRU(0 To 4) As String 'menu for mru
Dim savedas As String ' path to undo files
Dim Undobuffer() As String 'array of undo paths
Dim UndobufferCur() As Long 'array of cursor positions at undo
Dim UndoCount As Integer 'how many undos
Dim UndoPosition As Integer 'where we are in the undo/redo sequence




Private Sub Form_Load()
Dim x As Integer, temp As String
lastsaved = GetSetting(App.Title, "Paths", "Lastsaved", "C:\")
lastopened = GetSetting(App.Title, "Paths", "LastOpened", "C:\")
lastopenedEXT = GetSetting(App.Title, "Paths", "LastOpenedEXT", "txt")
For x = 0 To 4
    temp = GetSetting(App.Title, "Paths", "MRU" + Str(x), "")
    If temp <> "" Then
        mnuFileMRUSpace.Visible = True
        mnuFileMRU(x).Tag = temp
        mnuFileMRU(x).Caption = FileOnly(temp)
        mnuFileMRU(x).Visible = True
        MRU(x) = temp
    End If
Next x
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim x As Integer
SaveSetting App.Title, "Paths", "Lastsaved", lastsaved
SaveSetting App.Title, "Paths", "LastOpened", lastopened
SaveSetting App.Title, "Paths", "LastOpenedEXT", lastopenedEXT
For x = 0 To 4
    If mnuFileMRU(x).Visible Then
        SaveSetting App.Title, "Paths", "MRU" + Str(x), mnuFileMRU(x).Tag
    Else
        SaveSetting App.Title, "Paths", "MRU" + Str(x), ""
    End If
Next x
For x = 1 To UndoCount
    If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
Next x
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtftext.Top = TB.Top + TB.Height
rtftext.Height = Me.Height - rtftext.Top - 680
rtftext.Width = Me.Width - 90
End Sub

Private Sub mnuEditCopy_Click()
SendMessage rtftext.Hwnd, WM_COPY, 0, 0

End Sub

Private Sub mnuEditCut_Click()
SendMessage rtftext.Hwnd, WM_CUT, 0, 0
Backup
End Sub

Private Sub mnuEditDelete_Click()
SendMessage rtftext.Hwnd, WM_CLEAR, 0, 0
Backup

End Sub

Private Sub mnuEditPaste_Click()
SendMessage rtftext.Hwnd, WM_PASTE, 0, 0
Backup

End Sub

Private Sub mnuEditRedo_Click()
LockWindowUpdate rtftext.Hwnd
mnuEditUndo.Enabled = True
TB.Buttons(10).Enabled = True
rtftext.LoadFile Undobuffer(UndoPosition + 1)
rtftext.SelStart = UndobufferCur(UndoPosition + 1)
UndoPosition = UndoPosition + 1
If UndoPosition = UndoCount Then
    mnuEditRedo.Enabled = False
    TB.Buttons(11).Enabled = False
End If
LockWindowUpdate 0

End Sub

Private Sub mnuEditUndo_Click()
LockWindowUpdate rtftext.Hwnd
rtftext.LoadFile Undobuffer(UndoPosition - 1)
rtftext.SelStart = UndobufferCur(UndoPosition - 1)
If UndoPosition > 1 Then
    UndoPosition = UndoPosition - 1
End If
If UndoPosition < 2 Then
    mnuEditUndo.Enabled = False
    TB.Buttons(10).Enabled = False
End If
mnuEditRedo.Enabled = True
TB.Buttons(11).Enabled = True
LockWindowUpdate 0
End Sub


Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
Dim x As Integer
Dim isVis As Boolean
If FileExists(mnuFileMRU(Index).Tag) Then
    ClearUndo
    rtftext.LoadFile mnuFileMRU(Index).Tag
    curfile = mnuFileMRU(Index).Tag
    curEXT = ExtOnly(curfile)
    Backup
    mnuEditUndo.Enabled = False
    mnuEditRedo.Enabled = False
    TB.Buttons(10).Enabled = False
    TB.Buttons(11).Enabled = False
Else
    MsgBox "This file has been moved or deleted."
    MRU(Index) = ""
    mnuFileMRU(Index).Visible = False
    For x = 0 To mnuFileMRU.count - 1
        If mnuFileMRU(x).Visible = True Then isVis = True
    Next x
    mnuFileMRUSpace.Visible = isVis
End If

End Sub

Private Sub mnuFileNew_Click()
ClearUndo
rtftext.SelStart = 0
rtftext.SelLength = Len(rtftext.Text)
SendMessage rtftext.Hwnd, WM_CLEAR, 0, 0
Me.Caption = "Untitled"
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False

End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo woops
 With CommonDialog1
    If lastopenedEXT = "" Then lastopenedEXT = "txt"
    If lastopened = "" Then lastopened = "C:\"
    Select Case LCase(lastopenedEXT)
        Case "txt"
            .FilterIndex = 1
        Case "doc"
            .FilterIndex = 2
        Case Else
            .FilterIndex = 3
    End Select
    .initDir = lastopened
    .DialogTitle = "Open Text Files"
    .CancelError = True
    .Filter = "Text files (*.txt)|*.txt|Document files (*.doc)|*.doc|All files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    curfile = .FileName
    lastopened = PathOnly(.FileName)
    lastopenedEXT = ExtOnly(.FileName)
    ClearUndo
End With
curEXT = ExtOnly(curfile)
rtftext.LoadFile curfile
Me.Caption = FileOnly(curfile)
fixMRUs curfile
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False

woops: Exit Sub

End Sub

Private Sub mnuFileSave_Click()
Select Case LCase(curEXT)
    Case "txt"
        FileSave rtftext.Text, curfile
    Case "doc"
        rtftext.SaveFile curfile
    Case Else
        FileSave rtftext.Text, curfile
End Select
lastsaved = PathOnly(curfile)

End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo woops
If lastsaved = "" Then lastsaved = "C:\"
With CommonDialog1
    .initDir = lastsaved
    .DialogTitle = "Save Text Files"
    .CancelError = True
    .Filter = "Text files (*.txt)|*.txt|Document files (*.doc)|*.doc|All files (*.*)|*.*"
    Select Case LCase(curEXT)
        Case "txt"
            .FilterIndex = 1
        Case "doc"
            .FilterIndex = 2
        Case Else
            .FilterIndex = 3
    End Select
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    curfile = .FileName
    curEXT = ExtOnly(curfile)
    lastsaved = PathOnly(curfile)
End With
Select Case LCase(curEXT)
    Case "txt"
        FileSave rtftext.Text, curfile
    Case "doc"
        rtftext.SaveFile curfile
    Case Else
        FileSave rtftext.Text, curfile
End Select
Me.Caption = FileOnly(curfile)
woops: Exit Sub

End Sub
Public Function PathOnly(ByVal filepath As String) As String
Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
Dim temp As String
temp = Mid$(filepath, 1, InStrRev(filepath, "."))
temp = Left(temp, Len(temp) - 1)
If newext <> "" Then newext = "." + newext
ChangeExt = temp + newext
End Function


Private Sub rtftext_KeyUp(KeyCode As Integer, Shift As Integer)
Backup

End Sub

Private Sub rtftext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu mnuEdit
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "Undo"
        mnuEditUndo_Click
    Case "Redo"
        mnuEditRedo_Click
End Select
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


Public Sub fixMRUs(filepath As String)
Dim x As Integer, count As Integer
mnuFileMRU(0).Caption = FileOnly(filepath)
mnuFileMRU(0).Tag = filepath
mnuFileMRU(0).Visible = True
mnuFileMRUSpace.Visible = True
For x = 0 To 3
    If MRU(x) <> "" Then
        count = count + 1
        mnuFileMRU(count).Caption = FileOnly(MRU(x))
        mnuFileMRU(count).Tag = MRU(x)
        mnuFileMRU(count).Visible = True
    Else
        mnuFileMRU(x + 1).Visible = False
    End If
Next x
For x = 0 To 4
    If mnuFileMRU(x).Visible Then
        MRU(x) = mnuFileMRU(x).Tag
    End If
Next x

End Sub
Function FileExists(ByVal FileName As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
            End If
    End Select
End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function temppath() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    temppath = sBuffer
    If Right(temppath, 1) = "\" Then temppath = Left(temppath, Len(temppath) - 1)
End Function

Public Function GetTempFile(lpTempFilename As String) As Boolean
    lpTempFilename = String(255, vbNullChar)
    GetTempFile = GetTempFilename(temppath, "bb", 0, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function


Public Sub Backup()
Dim x As Integer
If UndoPosition <> UndoCount Then
    For x = UndoPosition + 1 To UndoCount
        If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
    Next x
    ReDim Preserve Undobuffer(UndoPosition)
    ReDim Preserve Undobuffer(UndoPosition + 10)
    UndoCount = UndoPosition
End If
If (UndoCount Mod 10) = 0 Then
    ReDim Preserve Undobuffer(UndoCount + 10)
    ReDim Preserve UndobufferCur(UndoCount + 10)
End If
UndoCount = UndoCount + 1
If GetTempFile(savedas) Then
    Undobuffer(UndoCount) = savedas
    UndobufferCur(UndoCount) = rtftext.SelStart
    rtftext.SaveFile savedas
End If
UndoPosition = UndoCount
mnuEditUndo.Enabled = True
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = mnuEditUndo.Enabled
TB.Buttons(11).Enabled = mnuEditRedo.Enabled

End Sub

Public Sub ClearUndo()
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False
Dim x As Integer
For x = 1 To UndoCount
    If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
Next x
ReDim Undobuffer(0 To 10)
UndoCount = 0
UndoPosition = 0

End Sub
