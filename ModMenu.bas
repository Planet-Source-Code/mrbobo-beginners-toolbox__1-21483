Attribute VB_Name = "ModMenu"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function GetMenu Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As bMENUINFO) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As bMENUINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewitem As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type bMENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type
Private Const MIM_APPLYTOSUBMENUS = &H80000000
Private Const MIM_BACKGROUND = &H2
Private Const LR_LOADFROMFILE = &H10
Private Const MF_BITMAP = &H4
Private Const MF_BYPOSITION = &H400
Public Sub SetMenuBackColor(mHwnd As Long, mColor As Long, mStyle As Long, mHatch As Long, SingleMnu As Boolean, Optional MmnuItem As Long)
'modified code from PSC
    Dim ret As Long, nCnt As Long, mnu As Integer
    Dim hMenu As Long, hSubMenu As Long
    Dim hBrush As Long
    Dim BI As LOGBRUSH
    Dim MI As bMENUINFO
    SetDefs mHwnd
    BI.lbStyle = mStyle
    BI.lbHatch = mHatch
    BI.lbColor = mColor
    hBrush = CreateBrushIndirect(BI)
    hMenu = GetMenu(mHwnd)
    nCnt = GetMenuItemCount(hMenu)
    If Not SingleMnu Then
        For mnu = 0 To nCnt - 1
            hSubMenu = GetSubMenu(hMenu, mnu)
            MI.cbSize = Len(MI)
            ret = GetMenuInfo(hSubMenu, MI)
            MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
            MI.hbrBack = hBrush
            ret = SetMenuInfo(hSubMenu, MI)
        Next mnu
    Else
        hSubMenu = GetSubMenu(hMenu, MmnuItem)
        MI.cbSize = Len(MI)
        ret = GetMenuInfo(hSubMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        ret = SetMenuInfo(hSubMenu, MI)
    End If
    If frmMenus.Check1.Value = 1 Then
        ret = GetMenuInfo(hMenu, MI)
        MI.fMask = MIM_BACKGROUND
        MI.hbrBack = hBrush
        ret = SetMenuInfo(hMenu, MI)
        DrawMenuBar mHwnd
    End If
End Sub

Public Sub SetDefs(mHwnd As Long)
'Sets things back to normal
    Dim ret As Long, nCnt As Long, mnu As Integer
    Dim hMenu As Long, hSubMenu As Long
    Dim hBrush As Long
    Dim MI As bMENUINFO
    hMenu = GetMenu(mHwnd)
    nCnt = GetMenuItemCount(hMenu)
    For mnu = 0 To nCnt - 1
        hSubMenu = GetSubMenu(hMenu, mnu)
        MI.cbSize = Len(MI)
        ret = GetMenuInfo(hSubMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        ret = SetMenuInfo(hSubMenu, MI)
    Next mnu
        ret = GetMenuInfo(hMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        ret = SetMenuInfo(hMenu, MI)
        DrawMenuBar mHwnd
End Sub

Public Sub SetMenuBitmap(mHwnd As Long, mImage As String, Mainmnu As Long, mSubmnu As Long)
'Not used in the example - fiddle with this one
'can work quite well if you choose the right pics
Dim hMenu As Long
Dim hSubMenu As Long
Dim lMenuID As Long
Dim hImage As Long
hMenu = GetMenu(mHwnd)
hSubMenu = GetSubMenu(hMenu, Mainmnu)
lMenuID = GetMenuItemID(hSubMenu, mSubmnu)
hImage = LoadImage(0, mImage, 0, 0, 0, LR_LOADFROMFILE)
ModifyMenu hSubMenu, mSubmnu, MF_BITMAP Or MF_BYPOSITION, lMenuID, hImage
End Sub

