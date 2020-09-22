Attribute VB_Name = "ListView32"
Option Explicit
' listview structures
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type LVHITTESTINFO
    pt As POINTAPI
    lFlags As Long
    lItem As Long
    lSubItem As Long
End Type
' send message to listview
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' styles for listview
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVM_FIRST = &H1000
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
' hittest constants
Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
' edit subitem constants
Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
Public Const LVIR_SELECTBOUNDS = 3
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Public Function ListView_HitTest(ListView As ListView, x As Single, y As Single) As LVHITTESTINFO
    On Error Resume Next
    Dim lRet As Long
    Dim lX As Long
    Dim lY As Long
    'x and y are in twips; convert them to pixels for the API call
    lX = x / Screen.TwipsPerPixelX
    lY = y / Screen.TwipsPerPixelY
    Dim tHitTest As LVHITTESTINFO
    With tHitTest
        .lFlags = 0
        .lItem = 0
        .lSubItem = 0
        .pt.x = lX
        .pt.y = lY
    End With
    'return the filled Structure to the routine
    lRet = SendMessage(ListView.hWnd, LVM_SUBITEMHITTEST, 0, tHitTest)
    ListView_HitTest = tHitTest
    Err.Clear
End Function
Public Sub ListView_ScaleEdit(LstView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)
    On Error Resume Next
    If tHitTest.lItem = -1 Then
        TextBox.Visible = False
        Err.Clear
        Exit Sub
    End If
    Dim XPixels As Integer
    Dim YPixels As Integer
    XPixels = Screen.TwipsPerPixelX
    YPixels = Screen.TwipsPerPixelY
    Dim tRec As RECT
    tRec.Top = tHitTest.lSubItem
    tRec.Left = LVIR_LABEL
    tRec.Bottom = 0
    tRec.Right = 0
    Dim lRet As Long
    lRet = SendMessage(LstView.hWnd, LVM_GETSUBITEMRECT, tHitTest.lItem, tRec)
    Dim lvRect As RECT
    lRet = GetClientRect(LstView.hWnd, lvRect)
    lvRect.Bottom = lvRect.Bottom * YPixels
    lvRect.Right = lvRect.Right * XPixels
    lvRect.Top = Round((LstView.Width - lvRect.Right) / 2)
    lvRect.Left = Round((LstView.Height - lvRect.Bottom) / 2)
    TextBox.Top = LstView.Top + lvRect.Top + tRec.Top * YPixels
    TextBox.Left = LstView.Left + lvRect.Left + tRec.Left * XPixels
    TextBox.Width = (tRec.Right - tRec.Left) * XPixels
    TextBox.Height = (tRec.Bottom - tRec.Top) * YPixels
    ' the scroll bar issue is complicated
    ' has to be treated individually, this has been through trial and error
    If ScrollBarVisible(LstView, SB_VERT) = True And ScrollBarVisible(LstView, SB_HORZ) = True Then
        ' if both scroll bars are available
        TextBox.Left = TextBox.Left - 110
        TextBox.Top = TextBox.Top - 90
        Err.Clear
        Exit Sub
    End If
    If ScrollBarVisible(LstView, SB_VERT) = True And ScrollBarVisible(LstView, SB_HORZ) = False Then
        TextBox.Top = TextBox.Top - 90
        Err.Clear
        Exit Sub
    End If
    If ScrollBarVisible(LstView, SB_VERT) = False And ScrollBarVisible(LstView, SB_HORZ) = True Then
        TextBox.Left = TextBox.Left - 110
        Err.Clear
        Exit Sub
    End If
    Err.Clear
End Sub
Public Sub ListView_BeforeEdit(ListView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)
    On Error Resume Next
    If tHitTest.lItem = -1 Then
        Err.Clear
        Exit Sub
    End If
    If tHitTest.lSubItem = 0 Then
        TextBox.Text = ListView.ListItems(tHitTest.lItem + 1).Text
    Else
        TextBox.Text = ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem)
    End If
    TextBox.Visible = True
    TextBox.SetFocus
    If Len(TextBox.Text) > 0 Then
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
    End If
    Err.Clear
End Sub
Public Sub ListView_AfterEdit(ListView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)
    On Error Resume Next
    Dim bEditMode As Boolean
    bEditMode = False
    If tHitTest.lItem > -1 Then
        If TextBox.Visible = True Then
            bEditMode = True
        End If
    End If
    TextBox.Visible = False
    If bEditMode = True Then
        If tHitTest.lSubItem = 0 Then
            ListView.ListItems(tHitTest.lItem + 1).Text = TextBox.Text
        Else
            ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem) = TextBox.Text
        End If
        tHitTest.lSubItem = (tHitTest.lSubItem + 1) Mod ListView.ColumnHeaders.Count
        If tHitTest.lSubItem = 0 Then
            tHitTest.lItem = (tHitTest.lItem + 1) Mod ListView.ListItems.Count
        End If
    End If
    Err.Clear
End Sub
Private Function ScrollBarVisible(LstView As ListView, ByVal fnBar As Long) As Boolean
    On Error Resume Next
    'returns true if lstreport's vertical scrollbar is visible
    Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo LstView.hWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
    Err.Clear
End Function
