VERSION 5.00
Begin VB.UserControl isExplorerBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   MouseIcon       =   "isExplorerBar.ctx":0000
   ScaleHeight     =   67
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   67
   ToolboxBitmap   =   "isExplorerBar.ctx":0152
   Begin VB.VScrollBar m_ScrollBar 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox m_pChild 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   -4500
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   -4500
      Top             =   4560
   End
   Begin VB.Image imgbuttons 
      Height          =   510
      Left            =   -1500
      Picture         =   "isExplorerBar.ctx":0464
      Top             =   6240
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "isExplorerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const strCurrentVersion = "1.91"
Private Const RDW_INVALIDATE = &H1
Private Const S_OK = 0
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const WM_USER = &H400
Private Const WM_THEMECHANGED       As Long = &H31A
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_CTLCOLORSCROLLBAR  As Long = &H137
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEHOVER         As Long = &H2A1
Private Const WM_SYSCOLORCHANGE     As Long = &H15 '21
Private Const TTS_NOPREFIX = &H2
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H0
Private Const COLOR_3DDKSHADOW As Long = 21
Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_BTNTEXT As Long = 18
Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Private Const COLOR_GRAYTEXT As Long = 17
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const COLOR_WINDOW As Long = 5
Private Const GWL_WNDPROC          As Long = -4
Private Const PATCH_05             As Long = 93                               'Table B (before) entry count
Private Const PATCH_09             As Long = 137                              'Table A (after) entry count
Private Type POINT
    x As Long
    y As Long
End Type
Private Type Size
    cx As Long
    cy As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type RGB            'Required for color trnsform using RGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Private Type TRIVERTEX          'For gradient Drawing
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Private Type TOOLINFO       'Tooltip Window Types
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type
Private Type BarItem        'Item Structure
    sParent As String       'Parent group key
    Key As String           'Key for external Access
    Index As Integer        'Index mainly for internal control
    Caption As String       '...you know
    Icon As Integer         'icon number.
    mRect As RECT           'rect in the control
    bOver As Boolean        'is the mouse over this?
    ToolTip As String        ' experimental - Anele Mbanga
    'iState As Integer       'Current State of the item
    Visible As Boolean
End Type
Private Type BarGroup
    Index As Integer        'Index for Internal Control'also external acces can be done with this, but It's easier to acces using the key than the index.
    Key As String           'Key for external access
    Type As Integer         'Experimental. I'll try to set each group as normal, details, Special or with child controls
        Caption As String       'Need more Information?:/
        Icon As Picture         'Group Icon
        items() As BarItem      'Array Of Items
        iItemsCount As Integer  'Count of Items in the group
        bExpanded As Boolean    'Is the group Expanded?
        mRect As RECT           'Rect of the group header
        bOver As Boolean        'Control variable, is the mouse over this?
        iState As Integer       'Current State of the group
        lItemsHeight As Long    'Group items frame height
        pChild As PictureBox    'Picture that act's as child for the group (Experimental)
        ChildHidden As Boolean  'Will indicate if child is hidden or not
        ToolTip As String       'experimental - Anele Mbanga
End Type
Private Type UxTheme        'Imported from a Cls File from VBAccelerator.com
    sClass As String        'And edited to keep the control in a single file.
    Part As Long            'I didn't used all the constant definitions where used
    State As Long           'in the original file, cuz I don't need them all
    hdc As Long             'But I added some others I need, like text offset
    hWnd As Long            'properties and UseTheme, to Detect If the draw was
    Left As Long            'succesfull or not, and then use classic windows Style
    Top As Long             'Drawing.
    Width As Long           'All the credits about the usage of UxTheme.dll defined on
    Height As Long          'cUxTheme.cls go for Steve at www.vbaccelerator.com
    Text As String
    TextAlign As DrawTextFlags
    IconIndex As Long
    hIml As Long
    RaiseError As Boolean
    UseThemeSize As Boolean
    UseTheme As Boolean
    TextOffset As Long
    RightTextOffset  As Long
End Type
Private Enum THEMESIZE
    'TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    'TS_DRAW            '// size that theme mgr will use to draw part
End Enum
Private Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum
Private Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum
Enum GRADIENT_FILL_RECT
    FillHor = GRADIENT_FILL_RECT_H
    FillVer = GRADIENT_FILL_RECT_V
End Enum
'Enum GRADIENT_TO_CORNER
'    All
'    TopLeft
'    TopRight
'    BottomLeft
'    BottomRight
'End Enum
'Enum CRADIENT_DIRECTION
'    DirectionSlash
'    DirectionBackSlash
'End Enum
Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum
Private Enum eMsgWhen
    MSG_AFTER = 1
    MSG_BEFORE = 2
    'MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum
#If False Then
Private MSG_AFTER, MSG_BEFORE, MSG_BEFORE_AND_AFTER
#End If
'*************************************************************
'
'   API Call Declares
'
'*************************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal hdc As Long, prc As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, prc As RECT, ByVal eSize As THEMESIZE, psz As Size) As Long
Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As DrawTextFlags, pBoundingRect As RECT, pExtentRect As RECT) As Long
Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, prcImage As RECT) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
'*************************************************************
'
'   Private Vars
'
'*************************************************************
'Subclassing Variables
Private nMsgCntB                   As Long                                    'Before msg table entry count
Private nMsgCntA                   As Long                                    'After msg table entry count
Private aMsgTblB()                 As Long                                    'Before msg table array
Private aMsgTblA()                 As Long                                    'After msg table array
Private nAddrSubclass              As Long                                    'The address of our WndProc
Private nAddrOriginal              As Long                                    'The address of the existing WndProc
Private sCode                      As String                                  'Binary subclass handler code string
'Control Variables
Private m_iTopOffset As Integer
Private m_cUxTheme As UxTheme
Private cGroups() As BarGroup
Private iGroups As Integer
Private m_objImageList As Object
Private iImgLType As Integer        'holds type of Imagelist
Private m_bOver As Boolean
Private m_NotOnUse As Long
Private m_GroupTextColor As OLE_COLOR
Private m_ItemTextColor As OLE_COLOR
Private m_ItemTextHoverColor As OLE_COLOR
Private m_GroupHoverColor As OLE_COLOR
Private m_bSpecialGroup As Boolean
Private m_SpecialGroup As BarGroup
Private m_SpecialGroupIcon As Picture
Private m_SpecialGroupBackground As Picture
Private m_bDetailsGroup As Boolean
Private m_DetailsGroup As BarGroup
Private m_DetailsGroupTittle As String
Private m_DetailsGroupText As String
Private m_DetailsRect As RECT
Private m_LastTextHeight As Long
Private m_Width As Long
Private m_AllowRedraw As Boolean
Private m_ttBackColor As Long 'properties for tooltip
Private m_ttTitle As String
Private m_ttForeColor As Long
Private m_ttIcon As ttIconType
Private m_ttlHwnd As Long
Private m_tti As TOOLINFO
Private bTrackMessages As Boolean
Private m_RedrawRect As RECT
Private m_DetailsPicture As StdPicture
Private sThemeFile As String
Private sColorName As String
Private UxThemeText As Boolean
Private bEnableVBAcIml As Boolean
'*************************************************************
'
'   Events Declares
'
'**************************************
Event MouseOver()
Event MouseOut()
Event GroupHover(sGroup As String)
Event GroupOut(sGroup As String)
Event ItemClick(sGroup As String, sItemKey As String)
Event GroupClick(sGroup As String, bExpanded As Boolean)
Event ItemHover(sGroup As String, sItemKey As String)
Event ItemOut(sGroup As String, sItemKey As String)
Public Enum GroupStatus
    Collapsed = 1
    Expanded = 2
End Enum
'*************************************************************
'
' Paul Caton Subclassing system.
'   a Huge work I have to thank him for.
'
'*************************************************************
'
'Subclass handler - MUST be the first Public routine in this file.
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lngHwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    On Error Resume Next
    'Parameters:
    'bBefore    Indicates whether the the message is being processed before or after the default handler - only really needed if a message is being subclassed before & after.
    'bHandled   Set this to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and optionaly, an 'after' callback
    'lReturn    Set this as per your intentions and requirements, see the MSDN documentation for each individual message type
    'hWnd       The window handle, should be the hWnd of the the User Control
    'uMsg       The message number
    'wParam     Message related data
    'lParam     Message related data
    'Notes:
    'If you really, really know what you're doing, it's possible to change the values
    'of the last four parameters in a 'before' callback so that different values get
    'passed to the default handler.. and optionaly, the 'after' callback
    Dim tmpval As Integer
    Select Case uMsg
    Case WM_CTLCOLORSCROLLBAR
        'Stop this message
        uMsg = 0
    Case WM_MOUSEWHEEL
        'Wheel movement.
        If m_ScrollBar.Visible Then
            If wParam = &H780000 Then
                'wparam contains the direction the wheel was moved.
                tmpval = m_ScrollBar.Value - 32
                m_ScrollBar.Value = IIf((tmpval < m_ScrollBar.Min), m_ScrollBar.Min, tmpval)
            ElseIf wParam = &HFF880000 Then
                tmpval = m_ScrollBar.Value + 32
                m_ScrollBar.Value = IIf((tmpval > m_ScrollBar.Max), m_ScrollBar.Max, tmpval)
            End If
        End If
    Case WM_MOUSELEAVE
    Case WM_MOUSEHOVER
    Case WM_MOUSEMOVE
    Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
        'Redraw!
        DoEvents
        UserControl_Paint
    End Select
    Err.Clear
End Sub
'======================================================================================================
'User Control's Subclass code
Private Sub Subclass_AddMsg(ByVal uMsg As Long, ByVal When As eMsgWhen)
    On Error Resume Next
    If When And eMsgWhen.MSG_BEFORE Then                                        'If Before
    'Add the message, pass the before table and before table message count variables ByRef
    Call zAddMsg(uMsg, aMsgTblB, nMsgCntB, eMsgWhen.MSG_BEFORE)
End If
If When And eMsgWhen.MSG_AFTER Then                                         'If After
'Add the message, pass the after table and after table message count variables ByRef
Call zAddMsg(uMsg, aMsgTblA, nMsgCntA, eMsgWhen.MSG_AFTER)
End If
Err.Clear
End Sub
'Delete the message from the msg table
Private Sub Subclass_DelMsg(ByVal uMsg As Long, ByVal When As eMsgWhen)
    On Error Resume Next
    If When And eMsgWhen.MSG_BEFORE Then                                        'If before
    'Delete the message, pass the Before table and before message count variables ByRef
    Call zDelMsg(uMsg, aMsgTblB, nMsgCntB, eMsgWhen.MSG_BEFORE)
End If
If When And eMsgWhen.MSG_AFTER Then                                         'If After
'Delete the message, pass the After table and after message count variables ByRef
Call zDelMsg(uMsg, aMsgTblA, nMsgCntA, eMsgWhen.MSG_AFTER)
End If
Err.Clear
End Sub
'Return whether we're running in the IDE. Public for general utility purposes
Private Function Subclass_InIDE() As Boolean
    On Error Resume Next
    Debug.Assert zSetTrue(Subclass_InIDE)
    Err.Clear
End Function
'Start the subclassing
Private Function Subclass_Start() As Boolean
    On Error Resume Next
    Const PATCH_01 As Long = 18                                                 'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02 As Long = 68                                                 'Address of the previous WndProc
    Const PATCH_03 As Long = 78                                                 'Relative address of SetWindowsLong
    Const PATCH_06 As Long = 116                                                'Address of the previous WndProc
    Const PATCH_07 As Long = 121                                                'Relative address of CallWindowProc
    Const PATCH_0A As Long = 186                                                'Address of the owner object
    Const FUNC_EBM As String = "EbMode"                                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"                                 'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const FUNC_CWP As String = "CallWindowProcA"                                'We use CallWindowProc to call the original WndProc
    Const MOD_VBA5 As String = "vba5"                                           'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                                           'Location of the EbMode function if running VB6
    Const MOD_USER As String = "user32"                                         'Location of the SetWindowLong  As Long CallWindowProc functions
    Dim i As Long                                                      'Loop index
    Dim sHex As String                                                    'Hex code string
    'Protect against double calling of Subclass_Start without having performed a Subclass_Stop first
    Debug.Assert (nAddrSubclass = 0)
    'Store the hex pair machine code representation in sHex
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE8xxxxx01x83F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168xxxxx02x6AFCFF7508E8xxxxx03xEBE031D24ABFxxxxx04xB9xxxxx05xE82D000000C3FF7514FF7510FF750CFF750868xxxxx06xE8xxxxx07x8945FCC331D2BFxxxxx08xB9xxxxx09xE801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B8xxxxx0Ax508B00FF90A4070000C3"
    'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
    For i = 1 To Len(sHex) Step 2                                               'For each pair of hex characters
        sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))                       'Convert a pair of hex characters to a byte and append to the ASCII string
    Next
    nAddrSubclass = StrPtr(sCode)                                               'Remember the address of the string code
    If Subclass_InIDE Then
        Call CopyMemory(ByVal nAddrSubclass + 15, &H9090, 2)                      'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
        i = zAddrFunc(MOD_VBA6, FUNC_EBM)                                         'Get the address of EbMode in vba6.dll
        If i = 0 Then                                                             'Found?
        i = zAddrFunc(MOD_VBA5, FUNC_EBM)                                       'VB5 perhaps, try vba5.dll
    End If
    Debug.Assert i                                                            'Ensure the EbMode function was found
    Call zPatchRel(PATCH_01, i)                                               'Patch the relative address to the EbMode api function
End If
nAddrOriginal = GetWindowLong(UserControl.hWnd, GWL_WNDPROC)                'Get the original window proc
Call zPatchVal(PATCH_02, nAddrOriginal)                                     'Original WndProc address for CallWindowProc, call the original WndProc
Call zPatchRel(PATCH_03, zAddrFunc(MOD_USER, FUNC_SWL))                     'Address of the SetWindowLong api function
Call zPatchVal(PATCH_05, 0)                                                 'Initial before table entry count
Call zPatchVal(PATCH_06, nAddrOriginal)                                     'Original WndProc address for SetWindowLong, unsubclass on IDE stop
Call zPatchRel(PATCH_07, zAddrFunc(MOD_USER, FUNC_CWP))                     'Address of the CallWindowProc api function
Call zPatchVal(PATCH_09, 0)                                                 'Initial after table entry count
Call zPatchVal(PATCH_0A, ObjPtr(Me))                                        'Get the address of the current instance of this User Control
nAddrOriginal = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, nAddrSubclass) 'Set our WndProc in place of the original
If nAddrOriginal <> 0 Then
    Subclass_Start = True                                                     'Success
End If
Debug.Assert Subclass_Start
Err.Clear
End Function
'Stop subclassing
Private Sub Subclass_Stop()
    On Error Resume Next
    Debug.Assert nAddrSubclass                                                  'Ensure that we are subclassing before we attempt to stop
    Call zPatchVal(PATCH_05, 0)                                                 'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(PATCH_09, 0)                                                 'Patch the Table A entry count to ensure no further 'after' callbacks
    Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, nAddrOriginal)            'Restore the original WndProc
    nMsgCntB = 0                                                                'Message before count set to zero
    nMsgCntA = 0                                                                'Message after count set to zero
    nAddrSubclass = 0                                                           'Indicate that we aren't subclassing
    Err.Clear
End Sub
'======================================================================================================
'These "z" routines are used by the subclass code - they shouldn't be called directly by the control author
'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen)
    On Error Resume Next
    Const PATCH_04 As Long = 88                                                 'Table B (before) address
    Const PATCH_08 As Long = 132                                                'Table A (after) address
    Dim nEntry As Long
    Dim nOff1 As Long
    Dim nOff2 As Long
    If uMsg = -1 Then                                                           'If all messages
    nMsgCnt = -1                                                              'Indicates that all messages shall callback
Else                                                                        'Else a specific message number
    For nEntry = 1 To nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
        Select Case aMsgTbl(nEntry)                                             'Select on the message number stored in this table entry
        Case -1                                                                 'This msg table slot is a deleted entry
            aMsgTbl(nEntry) = uMsg                                                'Re-use this entry
            Exit Sub                                                              'Bail
        Case uMsg                                                               'The msg is already in the table!
            Exit Sub                                                              'Bail
        End Select
    Next
    'Make space for the new entry
    ReDim Preserve aMsgTbl(1 To nEntry)                                       'Increase the size of the table. NB nEntry = nMsgCnt + 1
    nMsgCnt = nEntry                                                          'Bump the entry count
    aMsgTbl(nEntry) = uMsg                                                    'Store the message number in the table
End If
If When = eMsgWhen.MSG_BEFORE Then                                          'If before
nOff1 = PATCH_04                                                          'Offset to the Before table address
nOff2 = PATCH_05                                                          'Offset to the Before table entry count
Else                                                                        'Else after
nOff1 = PATCH_08                                                          'Offset to the After table address
nOff2 = PATCH_09                                                          'Offset to the After table entry count
End If
'Patch the appropriate table entries
Call zPatchVal(nOff1, zAddrMsgTbl(aMsgTbl))                                 'Patch the appropriate table address. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a message number is added.
Call zPatchVal(nOff2, nMsgCnt)                                              'Patch the appropriate table entry count
Err.Clear
End Sub
'Return the address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    On Error Resume Next
    zAddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
    'You may want to comment out the following line if you're using vb5 else the EbMode
    'GetProcAddress will stop here everytime because we look in vba6.dll first
    Debug.Assert zAddrFunc
    Err.Clear
End Function
'Return the address of the low bound of the passed table array
Private Function zAddrMsgTbl(ByRef aMsgTbl() As Long) As Long
    On Error Resume Next                                                        'The table may not be dimensioned yet so we need protection
    zAddrMsgTbl = VarPtr(aMsgTbl(1))                                            'Get the address of the first element of the passed message table
    Err.Clear
End Function
'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen)
    On Error Resume Next
    Dim nEntry As Long
    If uMsg = -1 Then                                                           'If deleting all messages
    nMsgCnt = 0                                                               'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                        'If before
    nEntry = PATCH_05                                                       'Patch the before table message count location
Else                                                                      'Else after
    nEntry = PATCH_09                                                       'Patch the after table message count location
End If
Call zPatchVal(nEntry, 0)                                                 'Patch the table message count
Else                                                                        'Else deleteting a specific message
For nEntry = 1 To nMsgCnt                                                 'For each table entry
    If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
    aMsgTbl(nEntry) = -1                                                  'Mark the table slot as available
    Exit For                                                              'Bail
End If
Next
End If
Err.Clear
End Sub
'Patch the machine code buffer offset with the relative address to the target address
Private Sub zPatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
    On Error Resume Next
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
    Err.Clear
End Sub
'Patch the machine code buffer offset with the passed value
Private Sub zPatchVal(ByVal nOffset As Long, ByVal nValue As Long)
    On Error Resume Next
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
    Err.Clear
End Sub
'Worker function for Subclass_InIDE - will only be called whilst running in the IDE
Private Function zSetTrue(bValue As Boolean) As Boolean
    On Error Resume Next
    zSetTrue = True
    bValue = True
    Err.Clear
End Function
'*************************************************************
'
'   User control Events
'
'*************************************************************
' Desc: Read the properties from the property bag -
'       also, a good place to start the subclassing
'       (if we're running) - this could also be enabled for
'       design time... if that's what you want.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With m_cUxTheme
        .hdc = UserControl.hdc
        .Width = 120
        .Height = 24
        .TextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
    End With
    m_ItemTextColor = 0
    m_GroupTextColor = 0
    m_NotOnUse = 0
    m_ItemTextHoverColor = RGB(127, 127, 127)
    m_GroupHoverColor = RGB(127, 127, 127)
    UserControl.Extender.Align = 3
    m_ttIcon = TTIconInfo
    m_ttTitle = App.Title
    With PropBag
        UserControl.FontName = .ReadProperty("FontName", UserControl.Ambient.Font.Name)
        UserControl.FontSize = .ReadProperty("FontSize", UserControl.Ambient.Font.Size)
        UserControl.Font.Charset = .ReadProperty("FontCharset")
        UxThemeText = CBool(.ReadProperty("UxThemeText", True))
        bEnableVBAcIml = CBool(.ReadProperty("EnableVBAcIml", False))
    End With
    'If we're not in design mode
    If Ambient.UserMode Then
        'Start subclassing
        Call Subclass_Start
        'Add the messages that we're interested in
        Call Subclass_AddMsg(WM_THEMECHANGED, MSG_AFTER)
        Call Subclass_AddMsg(WM_SYSCOLORCHANGE, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(WM_CTLCOLORSCROLLBAR, MSG_BEFORE)
        Call Subclass_AddMsg(WM_MOUSEWHEEL, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSELEAVE, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSEHOVER, MSG_AFTER)
    End If
    Err.Clear
End Sub
' Desc: Save the properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "FontName", UserControl.Font.Name, "Verdana"
        .WriteProperty "FontSize", UserControl.Font.Size, 8
        .WriteProperty "FontCharset", UserControl.Font.Charset
        .WriteProperty "UxThemeText", UxThemeText, True
    End With
    Err.Clear
End Sub
' Desc: Initialize control
Private Sub UserControl_Initialize()
    On Error Resume Next
    bEnableVBAcIml = False
    Err.Clear
End Sub
'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
    On Error Resume Next
    If nAddrSubclass <> 0 Then                                                  'If we're subclassing
    Call Subclass_Stop                                                        'Stop subclassing
End If
Err.Clear
End Sub
' Desc: when the scrollbar is visible and changes,
'   update the offset and redraw contents
Private Sub m_ScrollBar_Change()
    On Error Resume Next
    m_iTopOffset = m_ScrollBar.Value
    UserControl_Paint
    Err.Clear
End Sub
Private Sub m_pChild_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    UserControl_MouseMove 0, Shift, 3, 3
    Err.Clear
End Sub
Private Sub UserControl_AmbientChanged(PropertyName As String)
    On Error Resume Next
    UserControl_Paint
    Err.Clear
End Sub
' Desc: when the scrollbar is visible and changes,
'   update the offset and redraw contents
Private Sub m_ScrollBar_Scroll()
    On Error Resume Next
    m_ScrollBar_Change
    Err.Clear
End Sub
' Desc: to draw the apropiated background to the child
'       control, I'll try to caught that events here
Private Sub m_pChild_Paint(Index As Integer)
    On Error Resume Next
    'Child picturteboxes were redirected to this control
    'pChild(Index).hdc
    With m_cUxTheme
        If cGroups(Index).bExpanded Then
            If Not cGroups(Index).pChild Is Nothing Then
                'Child Picture Box Is Defined!
                Dim ltmpBackColor As Long
                cGroups(Index).pChild.Move cGroups(Index).mRect.Left * Screen.TwipsPerPixelX, (cGroups(Index).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(Index).mRect.Right - cGroups(Index).mRect.Left) * Screen.TwipsPerPixelX
                cGroups(Index).pChild.Visible = True
                cGroups(Index).pChild.AutoRedraw = True
                .hdc = cGroups(Index).pChild.hdc
                .hWnd = cGroups(Index).pChild.hWnd
                .Left = 0: .Top = 0: .Width = cGroups(Index).pChild.ScaleWidth: .Height = cGroups(Index).pChild.ScaleHeight
                .Part = 5
                .State = 1
                .Text = ""
                .Part = 5
                .State = 1
                Select Case sColorName
                Case "Metallic"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), vbWhite, B
                    ltmpBackColor = RGB(&HF0, &HF1, &HF5)
                Case "HomeStead"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), vbWhite, B
                    ltmpBackColor = RGB(&HF6, &HF6, &HEC)
                Case "Classic"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), GetSysColor(COLOR_WINDOW), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), GetSysColor(COLOR_BTNFACE), B
                    ltmpBackColor = GetSysColor(COLOR_WINDOW)
                Case Else
                    DrawTheme
                    ltmpBackColor = GetPixel(cGroups(Index).pChild.hdc, 4, 4) ' RGB(&HF0, &HF1, &HF5)
                End Select
                If Not .UseTheme Then
                    'Draw Failed, use Classic Style
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
                End If
                Dim tmpCtl As Variant
                For Each tmpCtl In UserControl.ParentControls
                    If tmpCtl.Container.Name = m_pChild(Index).Tag Then
                        If TypeName(tmpCtl) = "OptionButton" Then
                            tmpCtl.BackColor = ltmpBackColor
                        ElseIf TypeName(tmpCtl) = "Label" Then
                            tmpCtl.BackColor = ltmpBackColor
                        ElseIf TypeName(tmpCtl) = "Frame" Then
                            tmpCtl.BackColor = ltmpBackColor
                        ElseIf TypeName(tmpCtl) = "CheckBox" Then
                            tmpCtl.BackColor = ltmpBackColor
                        End If
                    End If
                Next
                .hdc = UserControl.hdc
                .hWnd = UserControl.hWnd
            Else
                'hide the child picturebox
                'cGroups(Index).pChild.Move cGroups(Index).mRect.Left * Screen.TwipsPerPixelX, (cGroups(Index).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(Index).mRect.Right - cGroups(Index).mRect.Left) * Screen.TwipsPerPixelX
                cGroups(Index).pChild.Visible = False
                'group has been drawn
            End If
        End If
    End With
    Err.Clear
End Sub
' desc: Here we process when the user Pushes
'       over items and header groups.
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Check for Click events
    'Process the current Events
    Dim ni As Integer
    Dim nj As Integer
    'Currently only left button actions supported.
    'If the button is different from vbleftbutton
    If Button = vbLeftButton Then
        'Check in the existing objects to see if anyone
        'has been presed
        If m_bSpecialGroup Then
            If y >= m_SpecialGroup.mRect.Top And y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
                'Mouse Down! Redraw Group Header
                m_SpecialGroup.iState = 3
                RedrawSpecialHeader
            End If
            If m_SpecialGroup.bExpanded Then
                'Analice each item for the group
                For nj = 1 To m_SpecialGroup.iItemsCount
                    'Search each item
                    If y >= m_SpecialGroup.items(nj).mRect.Top And y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
                        'Item down
                        RedrawItem -1, nj, 3
                    End If
                Next
            End If
        End If
        'Normal Groups
        For ni = 1 To iGroups
            If cGroups(ni).Key = "" Then GoTo NextGroup
            If y >= cGroups(ni).mRect.Top And y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
                'Mouse Down! Redraw Group Header
                cGroups(ni).iState = 3
                RedrawGroupHeader ni
            End If
            If cGroups(ni).bExpanded Then
                'Analice each item for the group
                For nj = 1 To cGroups(ni).iItemsCount
                    'Search each item
                    If y >= cGroups(ni).items(nj).mRect.Top And y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
                        'Item down
                        RedrawItem ni, nj, 3
                    End If
                Next
            End If
NextGroup:
        Next
        'Details Group
        If m_bDetailsGroup Then
            If y >= m_DetailsGroup.mRect.Top And y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
                'Mouse Down! Redraw Group Header
                m_DetailsGroup.iState = 3
                RedrawDetailsHeader
            End If
        End If
    End If
    Err.Clear
End Sub
' Desc: when the mouse pointer moves over the control,
'       some controls will be highlighted, other
'       deactivated. here we can process that events.
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Find out the Area where the mouse is located and highlight the current "object"
    Dim ni As Integer
    Dim nj As Integer
    'OldOver = m_bOver   'Set Previous State
    'm_bOver = (x > 0) And (y > 0) And (x < UserControl.ScaleWidth) And (y < UserControl.ScaleHeight)
    'm_bOver = (X > 0) And (Y > 0) And (X < UserControl.ScaleWidth - IIf(m_ScrollBar.Visible, m_ScrollBar.Width, 0)) And (Y < UserControl.ScaleHeight)
    'If (m_bOver And Not OldOver) Then
    '    RaiseEvent MouseOver
    'End If
    If UserControl.Enabled And Button = 0 Then
        If Not timUpdate.Enabled Then
            timUpdate.Enabled = True
            UserControl_MouseMove 0, 0, 1, 1
        Else
            timUpdate.Enabled = True
        End If
    End If
    DoEvents
    'Process the current Events
    'Start on the Special Group
    If m_bSpecialGroup Then
        If y >= m_SpecialGroup.mRect.Top And y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
            'Cursor is over the group
            If Not m_SpecialGroup.bOver Then
                m_SpecialGroup.bOver = True
                m_SpecialGroup.iState = 2
                SetHandCur True
                RedrawSpecialHeader
                'Raise Event for this group
                RaiseEvent GroupHover(m_SpecialGroup.Key)
                'added - anele mbanga
                If Len(m_SpecialGroup.ToolTip) > 0 Then
                    CreateTooltip m_SpecialGroup.Caption, m_SpecialGroup.ToolTip
                End If
                'm_ttTitle = "Tittle"
                'm_tti.lpStr = "Tooltip data"
            End If
        Else    'cursor is not over the group
            'Was In? then set out
            If m_SpecialGroup.bOver Then
                m_SpecialGroup.bOver = False
                m_SpecialGroup.iState = 1
                SetHandCur False
                RedrawSpecialHeader
                'Raise Event for this group
                RaiseEvent GroupOut(m_SpecialGroup.Key)
                ' added this section - Anele Mbanga
                If m_ttlHwnd <> 0 Then
                    DestroyWindow m_ttlHwnd
                End If
            End If
        End If
        If m_SpecialGroup.bExpanded Then
            'Analice each item for the group
            For nj = 1 To m_SpecialGroup.iItemsCount
                'Search each item
                If y >= m_SpecialGroup.items(nj).mRect.Top And y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
                    'Cursor Hover the item
                    If Not m_SpecialGroup.items(nj).bOver Then
                        'Set Hover
                        m_SpecialGroup.items(nj).bOver = True
                        RedrawItem -1, nj, 2
                        SetHandCur True
                        RaiseEvent ItemHover(m_SpecialGroup.Key, m_SpecialGroup.items(nj).Key)
                        'added - anele mbanga
                        If Len(m_SpecialGroup.items(nj).ToolTip) > 0 Then
                            CreateTooltip m_SpecialGroup.items(nj).Caption, m_SpecialGroup.items(nj).ToolTip
                        End If
                    End If
                Else
                    'Was Over this item?
                    If m_SpecialGroup.items(nj).bOver Then
                        'Set Out
                        m_SpecialGroup.items(nj).bOver = False
                        RedrawItem -1, nj, 1
                        SetHandCur False
                        RaiseEvent ItemOut(m_SpecialGroup.Key, m_SpecialGroup.items(nj).Key)
                        ' added this section - Anele Mbanga
                        If m_ttlHwnd <> 0 Then
                            DestroyWindow m_ttlHwnd
                        End If
                    End If
                End If
            Next
        End If
    End If
    ''Search in the normal groups
    For ni = 1 To iGroups
        If cGroups(ni).Key = "" Then GoTo NextGroup
        If y >= cGroups(ni).mRect.Top And y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
            'Cursor is over the group
            If Not cGroups(ni).bOver Then
                cGroups(ni).bOver = True
                cGroups(ni).iState = 2
                RedrawGroupHeader ni
                'Raise Event for this group
                SetHandCur True
                RaiseEvent GroupHover(cGroups(ni).Key)
                'added - anele mbanga
                If Len(cGroups(ni).ToolTip) > 0 Then
                    CreateTooltip cGroups(ni).Caption, cGroups(ni).ToolTip
                End If
            End If
        Else    'cursor is not over the group
            'Was In? then set out
            If cGroups(ni).bOver Then
                cGroups(ni).bOver = False
                cGroups(ni).iState = 1
                RedrawGroupHeader ni
                SetHandCur False
                'Raise Event for this group
                RaiseEvent GroupOut(cGroups(ni).Key)
                ' added this section - Anele Mbanga
                If m_ttlHwnd <> 0 Then
                    DestroyWindow m_ttlHwnd
                End If
            End If
        End If
        If cGroups(ni).bExpanded Then
            'Analice each item for the group
            For nj = 1 To cGroups(ni).iItemsCount
                'Search each item
                If y >= cGroups(ni).items(nj).mRect.Top And y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
                    'Cursor Hover the item
                    If Not cGroups(ni).items(nj).bOver Then
                        'Set Hover
                        cGroups(ni).items(nj).bOver = True
                        RedrawItem ni, nj, 2
                        SetHandCur True
                        'Raiseevent ItemOver
                        RaiseEvent ItemHover(cGroups(ni).Key, cGroups(ni).items(nj).Key)
                        'added - anele mbanga
                        If Len(cGroups(ni).items(nj).ToolTip) > 0 Then
                            CreateTooltip cGroups(ni).items(nj).Caption, cGroups(ni).items(nj).ToolTip
                        End If
                    End If
                Else
                    'Was Over this item?
                    If cGroups(ni).items(nj).bOver Then
                        'Set Out
                        cGroups(ni).items(nj).bOver = False
                        RedrawItem ni, nj, 1
                        SetHandCur False
                        'Raiseevent ItemOut
                        RaiseEvent ItemOut(cGroups(ni).Key, cGroups(ni).items(nj).Key)
                        ' added this section - Anele Mbanga
                        If m_ttlHwnd <> 0 Then
                            DestroyWindow m_ttlHwnd
                        End If
                    End If
                End If
            Next
        End If
NextGroup:
    Next
    'Search on the Details
    If m_bDetailsGroup Then
        If y >= m_DetailsGroup.mRect.Top And y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
            'Cursor is over the group
            If Not m_DetailsGroup.bOver Then
                m_DetailsGroup.bOver = True
                m_DetailsGroup.iState = 2
                SetHandCur True
                RedrawDetailsHeader
                'Raise Event for this group
                RaiseEvent GroupHover(m_DetailsGroup.Key)
            End If
        Else    'cursor is not over the group
            'Was In? then set out
            If m_DetailsGroup.bOver Then
                m_DetailsGroup.bOver = False
                m_DetailsGroup.iState = 1
                SetHandCur False
                RedrawDetailsHeader
                'Raise Event for this group
                RaiseEvent GroupOut(m_DetailsGroup.Key)
            End If
        End If
    End If
    Err.Clear
End Sub
' Desc: The clicks on the objects of the control,
'       are raised here, when the user releases
'       the button.
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ItemDoesntExist
    'Check for Click events
    'Process the current Events
    Dim ni As Integer
    Dim nj As Integer
    'Small Fix to allow clearStructure from an ItemClick event
    ' Thanks to Ademir Mazer Jr
    Dim GroupKeyAux As String
    Dim ItemKeyAux As String
    'Currently only left button actions supported.
    'If the button is different from vbleftbutton
    'then exit this sub.
    If Button = vbLeftButton Then
        'Search in special group
        If m_bSpecialGroup Then
            If y >= m_SpecialGroup.mRect.Top And y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
                'Cursor is over the group
                m_SpecialGroup.bExpanded = Not m_SpecialGroup.bExpanded
                m_SpecialGroup.iState = 2
                UserControl_Paint
                'SetHandCur True
                RaiseEvent GroupClick("-1", m_SpecialGroup.bExpanded)
            End If
            'Analice each item for the group
            If m_SpecialGroup.bExpanded Then
                For nj = 1 To m_SpecialGroup.iItemsCount
                    'Search each item
                    If y >= m_SpecialGroup.items(nj).mRect.Top And y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
                        'Cursor Hover the item
                        RedrawItem -1, nj, 2
                        'Small Fix to allow clearStructure from an ItemClick event
                        ' Thanks to Ademir Mazer Jr
                        GroupKeyAux = m_SpecialGroup.Key
                        ItemKeyAux = m_SpecialGroup.items(nj).Key
                        RaiseEvent ItemClick(GroupKeyAux, ItemKeyAux)
                    End If
                Next
            End If
        End If
        'Search the normal groups
        For ni = 1 To iGroups
            If cGroups(ni).Key = "" Then GoTo NextGroup
            If y >= cGroups(ni).mRect.Top And y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
                'Cursor is over the group
                cGroups(ni).bExpanded = Not cGroups(ni).bExpanded
                cGroups(ni).iState = 2
                UserControl_Paint
                UserControl.Refresh
                'SetHandCur True
                RaiseEvent GroupClick(cGroups(ni).Key, cGroups(ni).bExpanded)
            End If
            'Analice each item for the group
            If cGroups(ni).bExpanded Then
                For nj = 1 To cGroups(ni).iItemsCount
                    'Search each item
                    If y >= cGroups(ni).items(nj).mRect.Top And y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
                        'Cursor Hover the item
                        RedrawItem ni, nj, 2
                        GroupKeyAux = cGroups(ni).Key
                        ItemKeyAux = cGroups(ni).items(nj).Key
                        RaiseEvent ItemClick(GroupKeyAux, ItemKeyAux)
                    End If
                Next
            End If
NextGroup:
        Next
        'Search in Details group
        If m_bDetailsGroup Then
            If y >= m_DetailsGroup.mRect.Top And y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
                'Cursor is over the group
                m_DetailsGroup.bExpanded = Not m_DetailsGroup.bExpanded
                m_DetailsGroup.iState = 2
                UserControl_Paint
                UserControl.Refresh
                'SetHandCur True
                RaiseEvent GroupClick("-2", m_DetailsGroup.bExpanded)
            End If
        End If
    End If
ItemDoesntExist:
    Call UserControl_MouseMove(Button, Shift, x, y)
    Err.Clear
End Sub
' Desc: When the control is resized, redraw everything
Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl_Paint
    Err.Clear
End Sub
' Desc: This sub Is executed when the control Is shown
'       I added code to detect some messages that VB
'       don't notify.
Private Sub UserControl_Show()
    On Error Resume Next
    m_AllowRedraw = True
    UserControl_Paint
    '   please see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52168&lngWId=1
    '   Also thanks to Min Thant Sin at 7:33:33 AM on 5/3/2004
    '   see http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngTopicId=31065&lngWId=1&Forum=Visualbasic&TopicCategory=%20Request%20for%20Code
    '    'I used this to track some messages. But this feature generated
    '    '   too many bugs. So I quit. If you found this code usefull, you can use It.
    '    If UserControl.Ambient.UserMode Then
    '        bTrackMessages = True
    '        Do Until bTrackMessages = False
    '            DoEvents
    '            Call TrackMessage
    '            DoEvents
    '        Loop
    '    End If
    Err.Clear
End Sub
Public Sub Refresh()
    On Error Resume Next
    UserControl_Paint
    Err.Clear
End Sub
' Desc: This sub Is used to know If the mouse is
'       inside the control.
Private Sub timUpdate_Timer()
    On Error Resume Next
    'Check Out If the mouse is inside the control.
    timUpdate.Enabled = False
    If InBox(UserControl.hWnd) Then
        If m_bOver = False Then
            UserControl_Paint
            RaiseEvent MouseOver
        End If
        m_bOver = True
    Else
        If m_bOver Then
            'UserControl_Paint
            timUpdate.Enabled = False
            RaiseEvent MouseOut
        End If
        m_bOver = False
        'If any object was highlighted, reset all.
        UserControl_MouseMove 0, 0, 1, 1
    End If
    Err.Clear
End Sub
' Desc: This sub is where I draw the control objects.
'       Everything is here. maybe you can learn some
'       Things from here. I learned a lot from
'       VBThemeEplorer in vbaccelerator. the code for
'       drawing using UxTheme comes from that project.
'       but I turned the class into a structure and
'       changed his method Draw into a function (
'       Drawtheme), so, Now I don't need a extra file.
Private Sub UserControl_Paint()
    On Error Resume Next
    Dim ni As Integer
    Dim nj As Integer
    Dim iTop As Integer
    Dim tmpRect As RECT
    UserControl.Cls
    If Not m_AllowRedraw Then Exit Sub
    If Not UserControl.Ambient.UserMode Then
        'Stop filetring Messages
        bTrackMessages = False
        'Draw a Nice Banner :P
        With m_cUxTheme
            'Setup Some properties
            .hdc = UserControl.hdc
            .hWnd = UserControl.hWnd
            .sClass = "Explorerbar"
            .Part = 1
            .State = 1
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
            'Draw Background
            DrawTheme
            .Part = 9
            .Left = 3
            .Width = UserControl.ScaleWidth - 6
            .Height = 60
            .TextOffset = 0
            .RightTextOffset = 0
            .Top = 48
            .Text = "http://mx.geocities.com/fred_cpp/isexplorerbar"
            .TextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
            '.TextAlign = DT_CENTER Or DT_TOP Or DT_WORD_ELLIPSIS
            DrawTheme
            .Part = 12
            .Top = 25
            '.Left = 30
            '.Width = UserControl.ScaleWidth - 60
            .Height = 24
            .State = 2
            .TextAlign = DT_LEFT Or DT_CENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
            .Text = "isExplorerBar"
            DrawTheme
            'SetRect tmpRect, 10, .Top + 4, UserControl.Width - 10, .Top + 48
            'DrawRectText tmpRect, "http://mx.geocities.com/fred_cpp/isexplorerbar"
            If Not .UseTheme Then
                'No theme aviable, use classic drawing
                UserControl.Cls
                UserControl.Refresh
                SetRect tmpRect, 8, 12, UserControl.Width - 24, 34
                UserControl.Line (6, 12)-(UserControl.ScaleWidth - 12, 34), vbHighlight, BF
                UserControl.ForeColor = vbHighlightText
                UserControl.FontBold = True
                DrawText UserControl.hdc, "isExplorerBar", -1, tmpRect, DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_MODIFYSTRING
                SetRect tmpRect, 8, 35, UserControl.Width - 24, 88
                UserControl.Line (6, 35)-(UserControl.ScaleWidth - 12, 88), vbHighlight, B
                UserControl.ForeColor = vbButtonText
                UserControl.FontBold = False
                DrawText UserControl.hdc, "http://mx.geocities.com/fred_cpp/" & vbCrLf & "isexplorerbar", -1, tmpRect, DT_WORD_ELLIPSIS Or DT_MODIFYSTRING
            End If
            'PaintPicture toolboxbitmap?:( It's not possible? :/
        End With
    Else
        'Calculate the position and rects for each item.
        CalcRects
        'Get the theme name
        GetThemeName
        With m_cUxTheme
            'Setup Some properties
            .hdc = UserControl.hdc
            .hWnd = UserControl.hWnd
            .sClass = "Explorerbar"
            .Part = 1
            .State = 1
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
            .TextOffset = 32
            .RightTextOffset = 25
            .TextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
            Select Case sColorName
            Case "Metallic"
                'DoGradient RGB(&HC3, &HC7, &HD3), RGB(&HB1, &HB3, &HC8), FillVer, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
                'UserControl.BackColor = RGB(&HC3, &HC7, &HD3)
                'UserControl.Cls
                'DrawTheme
            Case "Classic"
                'No theme aviable, use classic drawing
                UserControl.Cls
                UserControl.Refresh
            Case Else
                'Other
                DrawTheme
            End Select
            'Check for the Special Group
            If m_bSpecialGroup Then
                'Draw the Special group
                RedrawSpecialHeader
                If m_SpecialGroup.bExpanded Then
                    'Draw the group Items frame
                    .Part = 9
                    .State = 1
                    .Text = ""
                    .Left = m_SpecialGroup.mRect.Left
                    .Top = m_SpecialGroup.mRect.Bottom
                    .Height = m_SpecialGroup.lItemsHeight
                    .Width = m_SpecialGroup.mRect.Right - m_SpecialGroup.mRect.Left
                    Select Case sColorName
                    Case "Metallic"
                        UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                        UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                    Case "HomeStead"
                        UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                        UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                    Case Else
                        DrawTheme
                    End Select
                    If Not .UseTheme Then
                        'Draw Failed, use Classic Style
                        UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbHighlight, B
                    End If
                    'Add back image
                    Dim dx As Integer
                    Dim dy As Integer
                    If Not m_SpecialGroupBackground Is Nothing Then
                        dx = m_SpecialGroupBackground.Width / Screen.TwipsPerPixelX
                        dy = m_SpecialGroupBackground.Height / Screen.TwipsPerPixelY
                        UserControl.ScaleMode = 3
                        UserControl.PaintPicture m_SpecialGroupBackground, .Left + 1, .Top + 1, .Width - 2, .Height - 2, , , , , vbSrcAnd
                    End If
                    'AlphaPaintPicture .Left + 1, .Top + 1, .Width - 2, .Height - 2, m_SpecialGroupBackground, 32
                    'Draw the items
                    For nj = 1 To m_SpecialGroup.iItemsCount
                        RedrawItem -1, nj, 0
                    Next
                    'group has been drawn
                    iTop = iTop + 6
                End If
                iTop = iTop + 6
            End If
            'for each group:
            For ni = 1 To iGroups
                'Draw Header
                If cGroups(ni).Key = "" Then GoTo NextGroup
                RedrawGroupHeader ni
                If cGroups(ni).bExpanded Then
                    If Not cGroups(ni).pChild Is Nothing Then
                        'Child Picture Box Is Defined!
                        cGroups(ni).pChild.Move cGroups(ni).mRect.Left * Screen.TwipsPerPixelX, (cGroups(ni).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(ni).mRect.Right - cGroups(ni).mRect.Left) * Screen.TwipsPerPixelX
                        cGroups(ni).pChild.Visible = True
                        .hdc = cGroups(ni).pChild.hdc
                        .hWnd = cGroups(ni).pChild.hWnd
                        .Left = 0: .Top = 0: .Width = cGroups(ni).pChild.ScaleWidth: .Height = cGroups(ni).pChild.ScaleHeight
                        .Part = 5
                        .State = 1
                        cGroups(ni).pChild.AutoRedraw = True
                        'Draw the group Items frame
                        .Text = ""
                        .Part = 5
                        .State = 1
                        m_pChild_Paint (ni)
                        'DrawTheme
                        .hdc = UserControl.hdc
                        .hWnd = UserControl.hWnd
                    Else
                        'Draw the group Items frame
                        .Top = cGroups(ni).mRect.Bottom
                        .Left = cGroups(ni).mRect.Left
                        .Height = cGroups(ni).lItemsHeight
                        .Width = cGroups(ni).mRect.Right - cGroups(ni).mRect.Left
                        .Text = ""
                        .Part = 5
                        .State = 1
                        Select Case sColorName
                        Case "Metallic"
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                        Case "HomeStead"
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                        Case Else
                            DrawTheme
                        End Select
                        If Not .UseTheme Then
                            'Draw Failed, use Classic Style
                            UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
                        End If
                        'Draw the items
                        For nj = 1 To cGroups(ni).iItemsCount
                            RedrawItem ni, nj, 0
                        Next
                        'group has been drawn
                    End If
                Else
                    'hide everything!
                    If Not cGroups(ni).pChild Is Nothing Then
                        'Child Picture Box Is Defined!
                        cGroups(ni).pChild.Visible = False
                    End If
                End If
NextGroup:
            Next
            'Details Group
            If m_bDetailsGroup Then
                ' Draw The Details Header
                RedrawDetailsHeader
                If m_DetailsGroup.bExpanded Then
                    'Draw the Tittle and text
                    .Part = 5
                    .State = 1
                    .Top = m_DetailsGroup.mRect.Bottom
                    .Left = m_DetailsGroup.mRect.Left
                    .Height = m_DetailsGroup.lItemsHeight
                    .Width = m_DetailsGroup.mRect.Right - m_DetailsGroup.mRect.Left
                    .Text = ""
                    Select Case sColorName
                    'this styles now are EMULATED. (just like microsoft does)
                Case "Metallic"
                    UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                    UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                Case "HomeStead"
                    UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF 'GetSysColor(COLOR_HIGHLIGHTTEXT), BF
                    UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                Case Else
                    DrawTheme
                End Select
                If Not .UseTheme Then
                    'Draw Failed, use Classic Style
                    UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
                End If
                'There Is a Image?
                If m_DetailsPicture Is Nothing Then
                    'No Image
                    'Draw Tittle
                    UserControl.FontUnderline = False
                    UserControl.FontBold = True
                    SetRect tmpRect, m_DetailsRect.Left, m_DetailsGroup.mRect.Bottom + 11, UserControl.ScaleWidth - 32, m_DetailsGroup.mRect.Bottom + 68
                    DrawText UserControl.hdc, m_DetailsGroupTittle, -1, tmpRect, DT_LEFT Or DT_WORDBREAK
                    'DrawText
                    UserControl.FontBold = False
                    DrawText UserControl.hdc, m_DetailsGroupText, -1, m_DetailsRect, DT_LEFT Or DT_WORDBREAK 'Len(m_DetailsGroupText)
                    RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
                    'group has been drawn
                Else
                    'We Have an Image move rects and go on
                    If m_DetailsPicture.Width > m_DetailsPicture.Height Then
                        'Calculate size
                        'wip
                    Else
                        'Calculate size again o_0
                        'wip
                    End If
                    'Draw Tittle
                    UserControl.FontUnderline = False
                    UserControl.FontBold = True
                    SetRect tmpRect, m_DetailsRect.Left, m_DetailsGroup.mRect.Bottom + 11 + UserControl.ScaleWidth - 128, UserControl.ScaleWidth - 128, m_DetailsGroup.mRect.Bottom + 11 + UserControl.ScaleWidth
                    DrawText UserControl.hdc, m_DetailsGroupTittle, -1, tmpRect, DT_LEFT Or DT_WORDBREAK
                    'DrawText
                    UserControl.FontBold = False
                    DrawText UserControl.hdc, m_DetailsGroupText, -1, m_DetailsRect, DT_LEFT Or DT_WORDBREAK 'Len(m_DetailsGroupText)
                    RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
                    'Draw Image
                    UserControl.PaintPicture m_DetailsPicture, 64, .Top + 8, UserControl.ScaleWidth - 128, UserControl.ScaleWidth - 128
                    'Draw Tittle
                    'group has been drawn
                End If
            End If
            iTop = iTop + 20
        End If
    End With
End If
UserControl.Refresh
Err.Clear
End Sub
'*************************************************************
'
'   Private Functions
'
'   Required Functions to make easier this ..thing
'
'**************************************
' Desc: On Version 0.9 and previous the rects of each item
'       where calulated on the Paint event of the usercontrol.
'       It Generated some problems, So I moved all that code
'       to a New Function. I Earned then almost 100 lines of
'       code!
Private Sub CalcRects()
    On Error Resume Next
    Dim ni As Integer
    Dim nj As Integer
    Dim iTop As Integer
    Dim ItemWidth As Long
    'Start variables
    iTop = -m_iTopOffset
    UserControl.FontBold = False
    'iTop = -m_iTopOffset
    m_Width = IIf(m_ScrollBar.Visible, UserControl.ScaleWidth - m_ScrollBar.Width, UserControl.ScaleWidth)
    'Check for the Special Group
    If m_bSpecialGroup Then
        'Set properties for the Special group
        iTop = iTop + 16    'Top Offset
        m_SpecialGroup.mRect.Top = iTop
        m_SpecialGroup.mRect.Left = 8
        m_SpecialGroup.mRect.Bottom = iTop + 24
        m_SpecialGroup.mRect.Right = m_Width - 8
        iTop = m_SpecialGroup.mRect.Bottom
        If m_SpecialGroup.bExpanded Then
            'Calculate Item's Rects
            iTop = iTop + 10
            For nj = 1 To m_SpecialGroup.iItemsCount
                m_SpecialGroup.items(nj).mRect.Top = iTop
                m_SpecialGroup.items(nj).mRect.Left = 20
                m_SpecialGroup.items(nj).mRect.Right = 40 + IIf(TextWidth((m_SpecialGroup.items(nj).Caption)) + 1 < (m_Width - 56), TextWidth((m_SpecialGroup.items(nj).Caption)) + 1, m_Width - 56)
                m_SpecialGroup.items(nj).mRect.Bottom = iTop + CalcHeightRectText(40, m_Width - 16, m_SpecialGroup.items(nj).Caption)
                iTop = m_SpecialGroup.items(nj).mRect.Bottom + 8
            Next
            m_SpecialGroup.lItemsHeight = iTop - m_SpecialGroup.mRect.Bottom + 8
            'group has been calculated
            iTop = iTop + 6
        End If
        iTop = iTop + 6
    End If
    'for each group:
    For ni = 1 To iGroups
        If cGroups(ni).Key = "" Then GoTo NextGroup
        'Calc Header Rect
        iTop = iTop + 10
        'Get Coordinates
        cGroups(ni).mRect.Top = iTop
        cGroups(ni).mRect.Left = 8
        cGroups(ni).mRect.Bottom = iTop + 24
        cGroups(ni).mRect.Right = m_Width - 8
        iTop = iTop + 24
        If cGroups(ni).bExpanded Then
            If Not cGroups(ni).pChild Is Nothing Then
                'Child Picture Box Is Defined!
                'Calculate the group Height
                iTop = iTop + cGroups(ni).pChild.ScaleHeight
                cGroups(ni).lItemsHeight = cGroups(ni).pChild.ScaleHeight
                'group has been Calculated
                iTop = iTop - 10
            Else
                iTop = iTop + 10
                'Calc the items
                For nj = 1 To cGroups(ni).iItemsCount
                    cGroups(ni).items(nj).mRect.Top = iTop
                    cGroups(ni).items(nj).mRect.Left = 20
                    cGroups(ni).items(nj).mRect.Right = 40 + IIf(TextWidth((cGroups(ni).items(nj).Caption)) + 1 < (m_Width - 56), TextWidth((cGroups(ni).items(nj).Caption)) + 1, m_Width - 56)
                    cGroups(ni).items(nj).mRect.Bottom = iTop + CalcHeightRectText(40, m_Width - 16, cGroups(ni).items(nj).Caption)
                    iTop = cGroups(ni).items(nj).mRect.Bottom + 8
                Next
                'Calculate the group Items frame
                cGroups(ni).lItemsHeight = iTop - cGroups(ni).mRect.Bottom + 12
                'group has been Calculated
                iTop = iTop + 6
            End If
        End If
        iTop = iTop + 12
NextGroup:
    Next
    'Details Group
    If m_bDetailsGroup Then
        iTop = iTop + 8
        'Get Coordinates
        m_DetailsGroup.mRect.Top = iTop
        m_DetailsGroup.mRect.Left = 8
        m_DetailsGroup.mRect.Bottom = iTop + 24
        m_DetailsGroup.mRect.Right = m_Width - 8
        iTop = m_DetailsGroup.mRect.Bottom
        Dim iTittleHeight As Integer
        If m_DetailsGroup.bExpanded Then
            'If there Is a Details Image...
            If m_DetailsPicture Is Nothing Then
                'There Isn't a Image
                UserControl.FontBold = True
                iTittleHeight = CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupTittle)
                UserControl.FontBold = False
                m_DetailsGroup.lItemsHeight = iTittleHeight + CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupText) + 24
                'Set the Details Rect
                UserControl.FontBold = True
                SetRect m_DetailsRect, 20, iTop + CalcHeightRectText(20, m_Width - 24, m_DetailsGroupTittle) + 12, m_Width - 24, iTop + 20 + m_DetailsGroup.lItemsHeight
                UserControl.FontBold = False
                iTop = m_DetailsRect.Bottom '+ 4
            Else
                'We Have An Image make room for It.
                iTop = iTop + 12 + UserControl.ScaleWidth - 128
                'Calculate the pos of the text and the tittle
                'Get the Height of the text
                UserControl.FontBold = True
                iTittleHeight = CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupTittle)
                UserControl.FontBold = False
                m_DetailsGroup.lItemsHeight = iTittleHeight + CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupText) + 24
                'Set the Details Rect
                UserControl.FontBold = True
                SetRect m_DetailsRect, 20, iTop + CalcHeightRectText(20, m_Width - 24, m_DetailsGroupTittle) + 12, m_Width - 24, iTop + 20 + m_DetailsGroup.lItemsHeight
                UserControl.FontBold = False
                iTop = m_DetailsRect.Bottom '+ 4
                m_DetailsGroup.lItemsHeight = iTop - m_DetailsGroup.mRect.Bottom - 12
            End If
            'group has been drawn
        End If
    End If
    'I'm re-using this variable, sorry,  Idon't want more variables on this sub.
    'this var should be called something like ScrollAmount
    'anyway, I think nobody will read this stuff:P If you do, thanks for look
    'into this code. Check out the Rect's array for each item in each group, I liked It a Lot!
    ItemWidth = iTop - UserControl.ScaleHeight + m_iTopOffset
    If ItemWidth = 0 Then
        'Setup ScrollBar
        'Adjust ScrollBar Properties
        m_ScrollBar.SmallChange = 4
        m_ScrollBar.LargeChange = UserControl.ScaleHeight
        m_ScrollBar.Max = 1 '(-ItemWidth) - 40
        m_ScrollBar.Move UserControl.ScaleWidth - m_ScrollBar.Width, 0, m_ScrollBar.Width, UserControl.ScaleHeight
        If m_ScrollBar.Visible = True Then
            m_ScrollBar.Visible = False
            CalcRects
        End If
        SetRect m_RedrawRect, 1, 1, UserControl.ScaleWidth - m_ScrollBar.Width - 2, UserControl.ScaleHeight - 2
        m_iTopOffset = 0
    ElseIf ItemWidth < 0 Then
        'Hide ScrollBar
        If m_ScrollBar.Visible Then
            m_ScrollBar.Visible = False
            m_iTopOffset = 0
            CalcRects
            SetRect m_RedrawRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
            Err.Clear
            Exit Sub
        Else
            m_ScrollBar.Visible = False
        End If
    Else
        'show and update scrollbar
        On Error GoTo NoHeight
        m_ScrollBar.SmallChange = 4
        m_ScrollBar.LargeChange = UserControl.ScaleHeight
        m_ScrollBar.Max = ItemWidth
        m_ScrollBar.Move UserControl.ScaleWidth - m_ScrollBar.Width, 0, m_ScrollBar.Width, UserControl.ScaleHeight
        SetRect m_RedrawRect, 0, 0, UserControl.ScaleWidth - m_ScrollBar.Width - 1, UserControl.ScaleHeight
        If Not m_ScrollBar.Visible Then
            'Prevent Infinite loop
            If Not UserControl.Extender.Visible Then Exit Sub
            'Scrollbar was not visible, recalculate rects, but before set to visible.
            m_ScrollBar.Visible = True
            DoEvents
            If m_AllowRedraw Then
                CalcRects
            End If
        End If
    End If
    Err.Clear
    Exit Sub
NoHeight:
    RaiseWarning "Couldn't Set ScrollBar Properties"
    Err.Clear
End Sub
' Desc: Calculate the height of a group box.
'       If there are multiline items, the
'       height won't be items*itemheight
Private Function CalcGroupHeight(iGroup As Integer) As Integer
    On Error Resume Next
    Dim nj As Integer
    Dim iTop As Integer
    Dim tmpHeight As Long           'USed to keep a copy oh m_LastTextHeight
    Dim textRect As RECT            'Copy used to calculate text height
    iTop = 24                       'Start up Offset
    tmpHeight = m_LastTextHeight    'Save var
    If iGroup = -1 Then
        With m_SpecialGroup
            For nj = 1 To .iItemsCount
                '.items(nj).mRect.Top = iTop
                SetRect textRect, .items(nj).mRect.Left + 20, .items(nj).mRect.Top, IIf((.items(nj).mRect.Right > m_Width - 12), m_Width - 12, .items(nj).mRect.Right), .items(nj).mRect.Bottom
                m_LastTextHeight = CalcHeightRectText(textRect.Left, textRect.Right, .items(nj).Caption)
                iTop = iTop + m_LastTextHeight + 8
            Next
        End With
        CalcGroupHeight = iTop
    Else    'Aplicar a grupo normal.
        With cGroups(iGroup)
            For nj = 1 To .iItemsCount
                'textRect.Top = iTop    'I don't know why I wrote this :/ ( Now I know, It's for looping o_0
                'Set the temp Rect
                SetRect textRect, .items(nj).mRect.Left + 20, .items(nj).mRect.Top, IIf((.items(nj).mRect.Right > m_Width - 12), m_Width - 12, .items(nj).mRect.Right), .items(nj).mRect.Bottom
                m_LastTextHeight = CalcHeightRectText(textRect.Left, textRect.Right, .items(nj).Caption)
                iTop = iTop + m_LastTextHeight + 8
            Next
        End With
        CalcGroupHeight = iTop
        'CalcGroupHeight = cGroups(iGroup).iItemsCount * 24 + 10
    End If
    m_LastTextHeight = tmpHeight    'Restore Var
    Err.Clear
End Function
' Desc: Draw Multilined text.
' Returns: Height of drawed Text
Private Sub DrawRectText(rtRect As RECT, sText As String)
    On Error Resume Next
    'draw text in the selected position
    m_LastTextHeight = CalcHeightRectText(rtRect.Left, rtRect.Right, sText)
    rtRect.Bottom = rtRect.Top + m_LastTextHeight
    DrawText UserControl.hdc, sText, Len(sText), rtRect, DT_LEFT Or DT_WORDBREAK
    'Redraw Window
    RedrawWindow UserControl.hWnd, rtRect, ByVal 0&, RDW_INVALIDATE
    Err.Clear
End Sub
' Desc: Draw Multilined text.
' Returns: Height of drawed Text
Private Function CalcHeightRectText(lLeft As Long, lright As Long, sText As String) As Long
    On Error Resume Next
    'Calculate vertical height of text Tittle + Text(wrapped)
    Dim rectText As RECT
    SetRect rectText, lLeft, 0, lright, UserControl.ScaleHeight
    CalcHeightRectText = DrawText(UserControl.hdc, sText, Len(sText), rectText, DT_CALCRECT Or DT_LEFT Or DT_WORDBREAK)
    Err.Clear
End Function
' Desc: Determine If the mouse cursor is inside a Object
Private Function InBox(ObjectHWnd As Long) As Boolean
    On Error Resume Next
    Dim mpos As POINT
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect ObjectHWnd, oRect
    If mpos.x >= oRect.Left And mpos.x <= oRect.Right And mpos.y >= oRect.Top And mpos.y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
    End If
    Err.Clear
End Function
'    on error resume next
'    'Use the API LineTo for Fast Drawing
'    UserControl.ForeColor = lColor
'    MoveToEx UserControl.hdc, X1, Y1, pt
'    LineTo UserControl.hdc, X2, Y2
'    Err.Clear
'End Sub
'Make Soft a color
Function SoftColor(lColor As OLE_COLOR) As OLE_COLOR
    On Error Resume Next
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lr As OLE_COLOR
    Dim lg As OLE_COLOR
    Dim lb As OLE_COLOR
    lr = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
    lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
    lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
    SoftColor = RGB(lr + lRed, lg + lGreen, lb + lBlue)
    Err.Clear
End Function
Private Function TranslateColor(origincolor As Long) As Long
    On Error Resume Next
    TranslateColor = OleTranslateColor(origincolor, 0, 0)
    Err.Clear
End Function
Function DoGradient(FromColor As Long, ToColor As Long, Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
    On Error Resume Next
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Long2RGB FromColor, R, G, B
    With vert(0)
        .x = Left
        .y = Top
        .Red = Val("&h" & Hex$(R) & "00")
        .Green = Val("&h" & Hex$(G) & "00")
        .Blue = Val("&h" & Hex$(B) & "00")
        .Alpha = 0&
    End With
    Long2RGB ToColor, R, G, B
    With vert(1)
        .x = Left + Width
        .y = Top + Height
        .Red = Val("&h" & Hex$(R) & "00")
        .Green = Val("&h" & Hex$(G) & "00")
        .Blue = Val("&h" & Hex$(B) & "00")
        .Alpha = 0&
    End With
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    DoGradient = GradientFillRect(UserControl.hdc, vert(0), 2, gRect, 1, DrawHorVer)
    Err.Clear
End Function
Sub Long2RGB(nColor As Long, Red As Byte, Green As Byte, Blue As Byte)
    On Error Resume Next
    Red = (nColor And &HFF&)
    Green = (nColor And &HFF00&) / &H100
    Blue = (nColor And &HFF0000) / &H10000
    Err.Clear
End Sub
'' Desc: a Alpha version of the paintpicture function
''       Still don't used
''Heavily based on this post:
''http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43879&lngWId=1
''with the suggestion on Dana Seaman and edited by me to make It a useful function
'    'The Structure will be replaced.
'    With BF
'        .BlendOp = AC_SRC_OVER
'        .BlendFlags = 0
'        .SourceConstantAlpha = lConstantAlpha
'        .AlphaFormat = 0
'    End With
'    'copy the BLENDFUNCTION-structure to a Long
'    RtlMoveMemory lBF, BF, 4
'
'    lBF = &H10000 * lConstantAlpha
'    m_tempImg.ScaleMode = 3
'    m_tempImg.Width = lPicture.Width / Screen.TwipsPerPixelX
'    m_tempImg.Height = lPicture.Height / Screen.TwipsPerPixelY
'    Set m_tempImg.Picture = lPicture
'    Set frmTest.Picture5.Picture = m_tempImg.Picture
'    'AlphaBlend
'    lr = AlphaBlend(UserControl.hdc, x, y, lwidth, lheight, m_tempImg.hdc, 0, 0, m_tempImg.ScaleWidth, m_tempImg.ScaleHeight, lBF)
'    If (lr = 0) Then
'       RaiseWarning Err.LastDllError
'    End If
'
'End Sub
' Desc: Convert a RGB color to long
Private Function RGBToLong(rgbColor As RGB) As Long
    On Error Resume Next
    RGBToLong = rgbColor.Blue + rgbColor.Green * 265 + rgbColor.Red * 65536
    Err.Clear
End Function
' Desc Convert a long into a RGB structure
Private Function LongToRGB(lColor As Long) As RGB
    On Error Resume Next
    LongToRGB.Red = lColor And &HFF
    LongToRGB.Green = (lColor \ &H100) And &HFF
    LongToRGB.Blue = (lColor \ &H10000) And &HFF
    Err.Clear
End Function
' Desc: This function will return  whether you are running
'       your program or DLL from within the IDE, or compiled.
Private Function InVBDesignEnvironment() As Boolean
    On Error Resume Next
    'Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11615&lngWId=1
    Dim strFileName As String
    Dim lngCount As Long
    strFileName = String$(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left$(strFileName, lngCount)
    InVBDesignEnvironment = False
    If UCase$(Right$(strFileName, 7)) = "VB5.EXE" Then
        InVBDesignEnvironment = True
    ElseIf UCase$(Right$(strFileName, 7)) = "VB6.EXE" Then
        InVBDesignEnvironment = True
    End If
    Err.Clear
End Function
' Desc: Get the Hand Cursor
Public Sub SetHandCur(Hand As Boolean)
    On Error Resume Next
    If Hand = True Then
        UserControl.MousePointer = 99
    Else
        UserControl.MousePointer = 0
    End If
    Err.Clear
End Sub
' Desc: Get a Group Index By Key
Function GetGroupsByKeyN(ByVal sGroupKey As Variant) As Integer
    On Error Resume Next
    Dim ni As Integer
    If (VarType(sGroupKey) <> vbInteger) And (VarType(sGroupKey) <> vbString) Then
        RaiseError "GetGroupsByKeyN: sGroupKey not of required Type (String or Integer)!"
        GetGroupsByKeyN = -3
        Err.Clear
        Exit Function
    End If
    'KEY was passed?
    If VarType(sGroupKey) = vbString Then
        'Check Normal Groups
        For ni = 1 To iGroups
            If LCase$(sGroupKey) = LCase$(cGroups(ni).Key) Then
                'this is the index
                GetGroupsByKeyN = ni
                Err.Clear
                Exit Function
            End If
        Next
        'Check Special Group
        If sGroupKey = "Special Group" Then
            GetGroupsByKeyN = -1
            Err.Clear
            Exit Function
            'Check Details Group
        ElseIf sGroupKey = "Details" Then
            GetGroupsByKeyN = -2
            Err.Clear
            Exit Function
            'Finally: String didn't match
        Else
            GetGroupsByKeyN = -3
            Err.Clear
            Exit Function
        End If
        'INDEX was passed
    Else
        GetGroupsByKeyN = sGroupKey
    End If
    Err.Clear
End Function
' Desc: Redraw a Single Item
Private Sub RedrawItem(iCurrentGroup As Integer, iItemNum As Integer, iState As Integer)
    On Error Resume Next
    Dim textRect As RECT
    'Set the text color
    Select Case iState
    Case 1  'Normal
        UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
    Case 2  'Hover
        UserControl.ForeColor = GetSysColor(COLOR_HIGHLIGHT)
    Case 3  'hot
        UserControl.ForeColor = GetSysColor(COLOR_GRADIENTACTIVECAPTION)
    Case 4  'Disabled
        UserControl.ForeColor = GetSysColor(COLOR_GRAYTEXT)
    End Select
    'Use underline style
    UserControl.FontUnderline = True
    UserControl.FontBold = False
    If iCurrentGroup = -1 Then
        With m_SpecialGroup.items(iItemNum)
            'Check for multiline text
            'if multiline text, adjust right
            'and adjust left to make room for the image
            SetRect textRect, .mRect.Left + 20, .mRect.Top, m_Width - 12, .mRect.Bottom
            DrawRectText textRect, .Caption
            On Error GoTo NoImage
            If iImgLType = 1 Then
                UserControl.PaintPicture m_objImageList.ListImages(.Icon).ExtractIcon, .mRect.Left, .mRect.Top, 16, 16
            ElseIf iImgLType = 2 Then
                m_objImageList.DrawImage .Icon, UserControl.hdc, .mRect.Left, .mRect.Top
            End If
        End With
    Else
        With cGroups(iCurrentGroup).items(iItemNum)
            'Set the rect where the text will be drawn
            SetRect textRect, .mRect.Left + 20, .mRect.Top, m_Width - 12, .mRect.Bottom
            'Draw the text
            DrawRectText textRect, .Caption
            On Error GoTo NoImage
            'Try to Draw the item image
            If iImgLType = 1 Then
                UserControl.PaintPicture m_objImageList.ListImages(.Icon).ExtractIcon, .mRect.Left, .mRect.Top, 16, 16
            ElseIf iImgLType = 2 Then
                m_objImageList.DrawImage .Icon, UserControl.hdc, .mRect.Left, .mRect.Top
            End If
        End With
    End If
    UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
    Err.Clear
    Exit Sub
NoImage:
    'No image or not imagelist was selected
    RaiseWarning "No Defined Imagelist or invalid Image Index"
    Err.Clear
End Sub
' Desc: Redraw a Group Header:
Private Sub RedrawGroupHeader(iCurrentGroup As Integer)
    On Error Resume Next
    Dim textRect As RECT
    Dim lcolor1 As Long
    Dim lcolor2 As Long
    'Setup Variables
    UserControl.FontUnderline = False
    UserControl.FontBold = True
    With cGroups(iCurrentGroup)
        m_cUxTheme.Part = 8
        m_cUxTheme.Left = .mRect.Left
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Width = .mRect.Right - .mRect.Left
        m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
        m_cUxTheme.State = .iState 'Now Support More States
        m_cUxTheme.Text = cGroups(iCurrentGroup).Caption
        m_cUxTheme.TextOffset = 0
        'Search for current theme and color scheme
        'Microsoft created the ExplorerBar with custom code and Images.
        'So We Need do somethig Similar. we will search for the theme file
        'and color Scheme
        Select Case sColorName
        Case "HomeStead"
            'this styles now are EMULATED. (just like microsoft does)
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
            SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_HIGHLIGHT), GetSysColor(COLOR_3DDKSHADOW))
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        Case "Metallic"
            'this styles now are EMULATED. (just like microsoft does)
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
            SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_3DDKSHADOW), GetSysColor(COLOR_BTNTEXT))
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        Case Else '"blue" and other themes
            DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbButtonFace, BF
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        End If
        'Draw Expand Button
        m_cUxTheme.Part = 7 + .bExpanded
        m_cUxTheme.Text = ""
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Left = m_Width - 32
        m_cUxTheme.Width = 24
        m_cUxTheme.Height = 24
        m_cUxTheme.State = .iState
        Select Case sColorName
        'this styles now are EMULATED. (just like microsoft does)
    Case "Metallic"
        UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
    Case "HomeStead"
        UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 18, 17, 17, vbSrcCopy
    Case Else
        DrawTheme
    End Select
    If Not m_cUxTheme.UseTheme Then
        'no theme aviable, use classic style
        If .iState = 3 Then  'pressed
        lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
    ElseIf .iState = 2 Then 'Hover
        lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
    Else    'Normal
        lcolor1 = vbButtonFace: lcolor2 = vbButtonFace
    End If
    'Draw Dutton
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
    UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
    'Draw arrow
    DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbButtonText
End If
RedrawWindow UserControl.hWnd, .mRect, ByVal 0&, RDW_INVALIDATE
End With
Err.Clear
End Sub
' Desc: Redraw a special Group Header
Private Sub RedrawSpecialHeader()
    On Error Resume Next
    Dim lcolor1 As Long
    Dim lcolor2 As Long
    Dim textRect As RECT
    UserControl.FontUnderline = False
    UserControl.FontBold = True
    With m_SpecialGroup
        m_cUxTheme.Part = 12
        m_cUxTheme.Left = .mRect.Left
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Width = .mRect.Right - .mRect.Left
        m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
        m_cUxTheme.State = .iState '(Doesn't support other states )
        m_cUxTheme.Text = .Caption
        m_cUxTheme.TextOffset = 36
        Select Case sColorName
        Case "Metallic"
            'this styles now are EMULATED. (just like microsoft does)
            DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
            DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
            DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
            SetRect textRect, .mRect.Left + m_cUxTheme.TextOffset, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_BTNHIGHLIGHT))
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        Case Else
            DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            SetRect textRect, .mRect.Left + m_cUxTheme.TextOffset, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbHighlight, BF
            UserControl.ForeColor = vbHighlightText
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbBlack
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        End If
        'm_cUxTheme.DrawThemeTextEx 1, iState
        'Draw Expand Button
        m_cUxTheme.TextOffset = 0
        m_cUxTheme.Part = 11 + .bExpanded
        m_cUxTheme.Text = ""
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Left = m_Width - 32
        m_cUxTheme.Width = 24
        m_cUxTheme.Height = 24
        m_cUxTheme.State = .iState
        Select Case sColorName
        Case "Metallic"
            UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
        Case Else
            DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            If .iState = 3 Then  'Pressed
            lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
        ElseIf .iState = 2 Then 'Hover
            lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
        Else    'normal
            lcolor1 = vbHighlight: lcolor2 = vbHighlight
        End If
        'Draw Dutton
        UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
        UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
        UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
        UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
        'Draw arrow
        DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbWindowBackground
    End If
    UserControl.PaintPicture m_SpecialGroupIcon, 12, .mRect.Top - 8, 32, 32 ', 0, 0, 32, 32
    RedrawWindow UserControl.hWnd, .mRect, ByVal 0&, RDW_INVALIDATE
    'm_LastTextHeight = .mRect.Bottom - .mRect.Top
End With
Err.Clear
End Sub
' Desc: Redraw the Details Group Header
Private Sub RedrawDetailsHeader()
    On Error Resume Next
    Dim textRect As RECT
    Dim lcolor1 As Long
    Dim lcolor2 As Long
    UserControl.FontUnderline = False
    UserControl.FontBold = True
    With m_DetailsGroup
        m_cUxTheme.Part = 8
        m_cUxTheme.Left = .mRect.Left
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Width = .mRect.Right - .mRect.Left
        m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
        m_cUxTheme.State = .iState '(Doesn't support other states )
        m_cUxTheme.Text = m_DetailsGroup.Caption
        m_cUxTheme.TextOffset = 0
        'Search for current theme and color scheme
        'Microsoft created the ExplorerBar with custom code and Images.
        'So We Need do somethig Similar. we will search for the theme file
        'and color Scheme
        Select Case sColorName
        Case "HomeStead"
            'this styles now are EMULATED. (just like microsoft does)
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
            DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
            SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_HIGHLIGHT), GetSysColor(COLOR_3DDKSHADOW))
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        Case "Metallic"
            'this styles now are EMULATED. (just like microsoft does)
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
            DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
            SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_3DDKSHADOW), GetSysColor(COLOR_BTNTEXT))
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        Case Else '"blue" and other themes
            DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            SetRect textRect, .mRect.Left + 4, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbButtonFace, BF
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hWnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbButtonText
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        End If
        'Draw Expand Button
        m_cUxTheme.Part = 7 + .bExpanded
        m_cUxTheme.State = .iState
        m_cUxTheme.Text = ""
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Left = m_Width - 32
        m_cUxTheme.Width = 24
        m_cUxTheme.Height = 24
        Select Case sColorName
        'this styles now are EMULATED. (just like microsoft does)
    Case "Metallic"
        UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
    Case "HomeStead"
        UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 18, 17, 17, vbSrcCopy
    Case Else
        DrawTheme
    End Select
    If Not m_cUxTheme.UseTheme Then
        'no theme aviable, use classic style
        If .iState = 3 Then  'Pressed
        lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
    ElseIf .iState = 2 Then 'Hover
        lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
    Else    'Normal
        lcolor1 = vbButtonFace: lcolor2 = vbButtonFace
    End If
    'Draw Dutton
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
    UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
    UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
    'Draw arrow
    DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbButtonText
End If
RedrawWindow UserControl.hWnd, .mRect, ByVal 0&, RDW_INVALIDATE
End With
Err.Clear
End Sub
' Desc: Draw the selected theme class, part, state on the especified rect
Private Function DrawTheme() As Boolean
    On Error Resume Next
    Dim hTheme As Long
    Dim bSuccess As Boolean
    Dim lr As Long
    Dim tTextR As RECT
    Dim tContentR As RECT
    Dim tImlR As RECT
    With m_cUxTheme
        If sColorName = "Classic" Then
            .UseTheme = False
            DrawTheme = False
            Err.Clear
            Exit Function
        End If
        bSuccess = True
        hTheme = OpenThemeData(.hWnd, StrPtr(.sClass))
        If (hTheme) Then
            'We Got an htheme
            .UseTheme = True
            Dim Tr As RECT
            Dim lWidthTaken As Long
            Tr.Left = .Left
            Tr.Top = .Top
            If (.IconIndex > -1) And (.hIml) Then
                ImageList_GetImageRect .hIml, .IconIndex, tImlR
                lWidthTaken = tImlR.Right - tImlR.Left + 4 + .TextOffset
            End If
            lWidthTaken = lWidthTaken + .TextOffset
            If (.UseThemeSize) Then
                Dim tSize As Size
                lr = GetThemePartSize(hTheme, .hdc, .Part, .State, Tr, TS_TRUE, tSize)
                Tr.Right = Tr.Left + tSize.cx
                Tr.Bottom = Tr.Top + tSize.cy
                lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, Tr, tContentR)
                If (.IconIndex > -1) And (.hIml) Then
                    If ((tContentR.Bottom - tContentR.Top) < (tImlR.Bottom - tImlR.Top + 4)) Then
                        Tr.Bottom = Tr.Bottom + ((tImlR.Bottom - tImlR.Top + 4) - (tContentR.Bottom - tContentR.Top))
                    End If
                    If ((tContentR.Right - tContentR.Left) < (tImlR.Right - tImlR.Left + 4)) Then
                        Tr.Right = Tr.Right + ((tImlR.Right - tImlR.Left + 4) - (tContentR.Right - tContentR.Left))
                    End If
                End If
                If Len(.Text) > 0 Then
                    lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, Tr, tContentR)
                    lr = GetThemeTextExtent(hTheme, .hdc, .Part, .State, StrPtr(.Text), -1, .TextAlign, Tr, tTextR)
                    If ((tContentR.Bottom - tContentR.Top) < (tTextR.Bottom - tTextR.Top)) Then
                        Tr.Bottom = Tr.Bottom + ((tTextR.Bottom - tTextR.Top) - (tContentR.Bottom - tContentR.Top))
                    End If
                    If ((tContentR.Right - tContentR.Left - lWidthTaken) < (tTextR.Right - tTextR.Left + 8)) Then
                        Tr.Right = Tr.Right + ((tTextR.Right - tTextR.Left + 8) - (tContentR.Right - tContentR.Left - lWidthTaken))
                    End If
                End If
            Else
                Tr.Right = .Left + .Width
                Tr.Bottom = .Top + .Height
            End If
            lr = DrawThemeParentBackground(.hWnd, .hdc, Tr)
            If (lr <> S_OK) Then
                bSuccess = False
                RaiseWarning "Failed to parent draw background for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
            End If
            lr = DrawThemeBackground(hTheme, .hdc, .Part, .State, Tr, Tr)
            If (lr <> S_OK) Then
                bSuccess = False
                'Important this is the main theme drawing procedure,
                'If this fail, then we can say the entire sub has
                'failed.
                .UseTheme = False
                RaiseWarning "Failed to draw background for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
            End If
            If Len(.Text) > 0 Then
                lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, Tr, tTextR)
                If (lr <> S_OK) Then
                    bSuccess = False
                    'RaiseWarning "Failed to retrieve background content rectangle for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
                End If
                tTextR.Left = tTextR.Left + lWidthTaken
                tTextR.Right = Tr.Right - .RightTextOffset
                tTextR.Top = Tr.Top
                tTextR.Bottom = Tr.Bottom
                If UxThemeText Then
                    'This will fail with far asian languages, replaced With custom DrawText
                    lr = DrawThemeText(hTheme, .hdc, .Part, .State, StrPtr(.Text), -1, .TextAlign, 0, tTextR)
                Else
                    Dim ltmpColor As Long
                    ltmpColor = UserControl.ForeColor
                    If .Part = 12 Then
                        UserControl.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
                    Else
                        UserControl.ForeColor = IIf(.State = 1, GetSysColor(COLOR_HIGHLIGHT), SoftColor(GetSysColor(COLOR_HIGHLIGHT)))
                    End If
                    DrawText .hdc, .Text, -1, tTextR, .TextAlign
                    UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
                End If
                If (lr <> S_OK) Then
                    bSuccess = False
                    'RaiseWarning "Failed to draw theme text for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
                End If
            End If
            If (.IconIndex > -1) Then
                Dim tIconR As RECT
                lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, Tr, tIconR)
                ImageList_GetImageRect .hIml, .IconIndex, tImlR
                tIconR.Left = tIconR.Left + 2
                tIconR.Top = tIconR.Top + 2
                tIconR.Right = tIconR.Left + tImlR.Right - tImlR.Left
                tIconR.Bottom = tIconR.Top + tImlR.Bottom - tImlR.Top
                lr = DrawThemeIcon(hTheme, .hdc, .Part, .State, tIconR, .hIml, .IconIndex)
                If (lr <> S_OK) Then
                    bSuccess = False
                    'RaiseWarning "Failed to draw theme icon for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
                End If
            End If
            CloseThemeData hTheme
            Dim tmpRect As RECT
            SetRect tmpRect, .Left, .Top, .Left + .Width, .Top + .Height
            RedrawWindow .hWnd, tmpRect, ByVal 0&, RDW_INVALIDATE
        Else
            RaiseWarning "No theme data for class '" & .sClass & "'.  - " & Err.LastDllError
            bSuccess = False
            .UseTheme = False
        End If
    End With
    DrawTheme = bSuccess
    Err.Clear
End Function
Private Sub GetThemeName()
    On Error Resume Next
    'Gett the current Theme name, ans Scheme Color
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim lPtrThemeFile As Long
    Dim lPtrColorName As Long
    Dim hres As Long
    Dim iPos As Long
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr("ExplorerBar"))
    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        hres = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If (iPos > 1) Then sThemeFile = Left$(sThemeFile, iPos - 1)
        sColorName = bColorName
        iPos = InStr(sColorName, vbNullChar)
        If (iPos > 1) Then sColorName = Left$(sColorName, iPos - 1)
        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid$(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left$(sThemeFile, iPos)
                Exit For
            End If
        Next
        sShellStyle = sShellStyle & "Shell\" & sColorName & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        sColorName = "Classic"
    End If
    Err.Clear
End Sub
' Desc: This small sub draws the arrow in the selected position
Private Sub DrawArrow(ByVal x As Integer, ByVal y As Integer, ByVal bUp As Boolean, ByVal lColor As Long)
    On Error Resume Next
    If bUp Then
        UserControl.Line (x + 9, y + 11)-(x + 13, y + 7), lColor
        UserControl.Line (x + 10, y + 11)-(x + 13, y + 8), lColor
        UserControl.Line (x + 15, y + 11)-(x + 11, y + 7), lColor
        UserControl.Line (x + 14, y + 11)-(x + 11, y + 8), lColor
        UserControl.Line (x + 9, y + 15)-(x + 13, y + 11), lColor
        UserControl.Line (x + 10, y + 15)-(x + 13, y + 12), lColor
        UserControl.Line (x + 15, y + 15)-(x + 11, y + 11), lColor
        UserControl.Line (x + 14, y + 15)-(x + 11, y + 12), lColor
    Else
        UserControl.Line (x + 9, y + 8)-(x + 13, y + 12), lColor
        UserControl.Line (x + 10, y + 8)-(x + 13, y + 11), lColor
        UserControl.Line (x + 15, y + 8)-(x + 11, y + 12), lColor
        UserControl.Line (x + 14, y + 8)-(x + 11, y + 11), lColor
        UserControl.Line (x + 9, y + 12)-(x + 13, y + 16), lColor
        UserControl.Line (x + 10, y + 12)-(x + 13, y + 15), lColor
        UserControl.Line (x + 15, y + 12)-(x + 11, y + 16), lColor
        UserControl.Line (x + 14, y + 12)-(x + 11, y + 15), lColor
    End If
    Err.Clear
End Sub
' Desc: Show an Error to the programmer Is using the control
Private Sub RaiseError(sErrorDescription As String)
    On Error Resume Next
    MsgBox "An Error has occurred!" & vbCrLf & sErrorDescription, vbCritical, "isExplorerBar"
    Err.Clear
End Sub
' Desc: Show Warning in the Debug Window
Private Sub RaiseWarning(sWarning As String)
    On Error Resume Next
    Err.Clear
End Sub
' Desc: Create tooltip!
' This is still a work in progress function
Private Sub CreateTooltip(sTittle As String, sCaption As String)
    On Error Resume Next
    Dim lpRect As RECT
    Dim lWinStyle As Long
    ' added this section - Anele Mbanga
    m_ttTitle = Replace$(sTittle, "&&", "&")
    m_tti.lpStr = sCaption
    ' remove comment on this block - anele mbanga
    If m_ttlHwnd <> 0 Then
        DestroyWindow m_ttlHwnd
    End If
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    ''create baloon style if desired
    'If mvarStyle = TTBalloon Then
    lWinStyle = lWinStyle Or TTS_BALLOON
    ''the parent control has to have been set first
    'If Not mvarParentControl Is Nothing Then
    m_ttlHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
    ''make our tooltip window a topmost window
    SetWindowPos m_ttlHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    ''get the rect of the parent control
    GetClientRect UserControl.hWnd, lpRect
    ''now set our tooltip info structure
    With m_tti
        ''if we want it centered, then set that flag
        'If mvarCentered Then
        .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
        'Else
        .lFlags = TTF_SUBCLASS
        'End If
        ''set the hwnd prop to our parent control's hwnd
        .lHwnd = UserControl.hWnd
        .lId = 0
        .hInstance = App.hInstance
        '.lpstr = ALREADY SET
        .lpRect = lpRect
    End With
    ''add the tooltip structure
    SendMessage m_ttlHwnd, TTM_ADDTOOLA, 0&, m_tti
    ''if we want a title or we want an icon
    If m_ttTitle <> vbNullString Or m_ttIcon <> TTNoIcon Then
        SendMessage m_ttlHwnd, TTM_SETTITLE, CLng(m_ttIcon), ByVal m_ttTitle
    End If
    If m_ttForeColor <> Empty Then
        SendMessage m_ttlHwnd, TTM_SETTIPTEXTCOLOR, m_ttForeColor, 0&
    End If
    If m_ttBackColor <> Empty Then
        SendMessage m_ttlHwnd, TTM_SETTIPBKCOLOR, m_ttBackColor, 0&
    End If
    'End If
    Err.Clear
End Sub
'**************************************************************
'Function:      GetItemIndex / Private
'Description:   Returns the index of an BarItem, passed as variant
'               Item can only be String or Integer here
'Parameters:    selGroup:   BarGroup, the group we search in
'               Item:       Variant, containing Items parameter
'Result:        0 for no item found
'               Items index for success
'**************************************************************
Private Function GetItemIndex(selgroup As BarGroup, Item As Variant) As Integer
    On Error Resume Next
    Dim nj As Integer
    'Check if there are Items in the group
    If selgroup.iItemsCount > 0 Then
        'First check the VarType of Item
        'STRING
        If VarType(Item) = vbString Then
            For nj = 1 To selgroup.iItemsCount
                If LCase$(selgroup.items(nj).Key) = LCase$(Item) Then
                    GetItemIndex = nj
                    Err.Clear
                    Exit Function
                End If
            Next
            'When we get here, there is no Item with this key
            'RaiseError "GetItemIndex/String: Item specified not found!"
            Err.Clear
            Exit Function
            'INTEGER
        ElseIf (VarType(Item) = vbInteger) Then
            'Does this Item Index exist?
            If (Item >= 1) And (Item <= selgroup.iItemsCount) Then
                GetItemIndex = Item
                Err.Clear
                Exit Function
            Else
                'RaiseError "GetItemIndex/Integer: Item specified not found!"
                GetItemIndex = 0
                Err.Clear
                Exit Function
            End If
        Else
            'RaiseError "GetItemIndex: Item must contain String or Integer!"
            GetItemIndex = 0
            Err.Clear
            Exit Function
        End If
        'when we get here, there is no item in this group
    Else
        'RaiseError "GetItemIndex: There are no Items in this group!"
        GetItemIndex = 0
        Err.Clear
        Exit Function
    End If
    'and when we get here, something else went wrong
    'RaiseError "GetItemIndex: Unknown error!"
    GetItemIndex = 0
    Err.Clear
End Function
'**************************************************************
'Function:      GroupExists / Private
'Description:   Checks if the specified Group exists
'Parameters:    Index:      Integer, the group's index we want
'                           to check
'Result:        False:      Group dosn't exist
'               True:       Group exists
'**************************************************************
Public Function DoesGroupExists(Group) As Boolean
    On Error Resume Next
    Dim iGroupIndex As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    DoesGroupExists = GroupExists(iGroupIndex)
    Err.Clear
End Function
Private Function GroupExists(Index As Integer) As Boolean
    On Error GoTo GroupError
    Dim dummy As String
    Select Case Index
    Case Is > 0
        dummy = cGroups(Index).Key
        GroupExists = True
        Err.Clear
        Exit Function
    Case -1
        dummy = m_SpecialGroup.Key
        GroupExists = True
        Err.Clear
        Exit Function
    Case -2
        dummy = m_DetailsGroup.Key
        GroupExists = True
        Err.Clear
        Exit Function
    Case Else
        GroupExists = False
        Err.Clear
        Exit Function
    End Select
GroupError:
    GroupExists = False
    Err.Clear
End Function
'**************************************************************
'Function:      GetIconIndex / Private
'Description:   Returns the index of an image from ImageList,
'               passed as variant (Index or Key)
'               iIcon can only be String or Integer here
'Parameters:    iIcon:   Key or Index of Imagelist
'Result:        -1 for no icon found
'               Icon index for success
'**************************************************************
Private Function GetIconIndex(iIcon As Variant) As Integer
    On Error GoTo NoImage
    Dim i As Integer
    Dim iLCnt As Integer
    'Parameter NOT string or integer?
    If (VarType(iIcon) <> vbInteger) And (VarType(iIcon) <> vbString) Then
        RaiseError "GetIconIndex: iIcon not of required Type (String or Integer)!"
        GetIconIndex = -1
        Err.Clear
        Exit Function
    End If
    If iImgLType = 1 Then
        iLCnt = m_objImageList.ListImages.Count
    ElseIf iImgLType = 2 Then
        iLCnt = m_objImageList.ImageCount
    End If
    'Key was passed
    If VarType(iIcon) = vbString Then
        'get icon index
        For i = 1 To iLCnt
            If LCase$(m_objImageList.ListImages(i).Key) = LCase$(iIcon) Then
                'we did find the Icons index
                GetIconIndex = i
                Err.Clear
                Exit Function
            End If
        Next
        'when we got here the string doesn't match
        RaiseError "GetIconIndex: icon with key " & iIcon & " doesn't exist!"
        GetIconIndex = -1
        Err.Clear
        Exit Function
    End If
    'Index was passed
    If iIcon >= 1 Or iIcon <= iLCnt Then
        GetIconIndex = iIcon
    Else
        RaiseWarning "GetIconIndex: invalid Image Index!"
        GetIconIndex = -1
    End If
    Err.Clear
    Exit Function
NoImage:
    'No imagelist was selected
    RaiseWarning "No Defined Imagelist"
    GetIconIndex = -1
    Err.Clear
End Function
'*************************************************************
'
'   Public Functions
'
'   I'll try to add each element in runtime. I'll provide
'   all the needed functions Add groups, add items, clear,
'   and a event response for a click on each element
'
'**************************************
' Desc: Add a Group to the control
' Some parameters Still don't work, cuz I'm implementing changes.
Public Sub AddGroup(sKey As String, sCaption As String, Optional ToolTip As String = "", Optional State As GroupStatus = Collapsed, Optional MyChild As Picture)
    On Error Resume Next
    m_NotOnUse = 1
    iGroups = iGroups + 1
    ReDim Preserve cGroups(iGroups)
    With cGroups(iGroups)
        .Caption = sCaption
        .Key = LCase$(sKey)
        '.Icon = iIcon
        .bExpanded = True
        If IsMissing(ToolTip) = False Then .ToolTip = ToolTip
        If IsMissing(MyChild) = False Then
            SetGroupChild sKey, MyChild
        End If
        If State = Collapsed Then
            ExpandGroup sKey, False
        Else
            ExpandGroup sKey, True
        End If
    End With
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Add a Item to a group in the the control
Public Sub AddItem(Group, sKey As String, sCaption As String, Optional iIcon As Variant, Optional ToolTip As String, Optional ItemToHide As String = "")                'Integer)
    On Error Resume Next
    Dim iCurrentGroup As Integer
    If MvSearch(ItemToHide, sKey, ",") >= 0 Then Exit Sub
    If Not IsMissing(iIcon) Then
        iIcon = GetIconIndex(iIcon)
    Else
        iIcon = -1
    End If
    iCurrentGroup = GetGroupsByKeyN(Group)
    m_NotOnUse = 1
    If iCurrentGroup = -1 Then
        m_SpecialGroup.iItemsCount = m_SpecialGroup.iItemsCount + 1 'Get Current count (+1)
        ReDim Preserve m_SpecialGroup.items(m_SpecialGroup.iItemsCount) 'Redim array
        With m_SpecialGroup.items(m_SpecialGroup.iItemsCount)
            .Key = LCase$(sKey)
            .Caption = sCaption
            .sParent = "Special Group"
            .Index = m_SpecialGroup.iItemsCount
            .Icon = iIcon
            ' added line - Anele Mbanga
            If IsMissing(ToolTip) = False Then .ToolTip = ToolTip
        End With
    Else
        If iCurrentGroup = -3 Then
            RaiseWarning "Can't assign items to the Specified group"
            Err.Clear
            Exit Sub
        End If
        If iCurrentGroup = 0 Then GoTo noSuchGroup
        cGroups(iCurrentGroup).iItemsCount = cGroups(iCurrentGroup).iItemsCount + 1 'Get Current count (+1)
        ReDim Preserve cGroups(iCurrentGroup).items(cGroups(iCurrentGroup).iItemsCount) 'Redim array
        With cGroups(iCurrentGroup).items(cGroups(iCurrentGroup).iItemsCount)
            .Key = LCase$(sKey)
            .Caption = sCaption
            .sParent = Group
            .Index = cGroups(iCurrentGroup).iItemsCount
            .Icon = iIcon
            ' added line - Anele Mbanga
            If IsMissing(ToolTip) = False Then .ToolTip = ToolTip
        End With
    End If
    UserControl_Paint
    Err.Clear
    Exit Sub
noSuchGroup:
    RaiseWarning "The group '" & Group & "' doesn't exist"
    Err.Clear
End Sub
' Desc: Set the image list object where we get the icons.
Public Sub SetImageList(ByRef ImageListObj As Object)
    On Error Resume Next
    Set m_objImageList = ImageListObj
    '**********************************
    If TypeName(m_objImageList) = "ImageList" Then
        iImgLType = 1
    ElseIf TypeName(ImageListObj) = "vbalImageList" Then
        iImgLType = 2
    Else
        iImgLType = 0
        'its possible to raise an error here but not really needed?
    End If
    '**********************************
    Err.Clear
End Sub
' Desc: Set Up the Special Group (there is only a special group in each control)
Public Sub AddSpecialGroup(Caption As String, Optional Icon As Picture, Optional background As Picture, Optional ToolTip As String)
    On Error Resume Next
    m_bSpecialGroup = True
    m_SpecialGroup.Caption = Caption
    m_SpecialGroup.Key = "Special Group"
    m_SpecialGroup.bExpanded = True
    If IsMissing(ToolTip) = False Then m_SpecialGroup.ToolTip = ToolTip
    m_NotOnUse = 1
    Set m_SpecialGroupIcon = Icon
    Set m_SpecialGroupBackground = background
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Hide the special group
Public Sub HideSpecialGroup()
    On Error Resume Next
    m_bSpecialGroup = False
    UserControl_Paint
    UserControl.Refresh
    Err.Clear
End Sub
'Setup the Details Group in the control.
Public Sub AddDetailsGroup(Caption As String, sDetails As String, Optional sTittle As String = "")
    On Error Resume Next
    m_NotOnUse = 1
    m_bDetailsGroup = True
    m_DetailsGroup.Caption = Caption
    m_DetailsGroup.Key = "Details Group"
    m_DetailsGroup.Caption = Caption
    m_DetailsGroupTittle = sTittle
    m_DetailsGroupText = sDetails
    m_DetailsGroup.bExpanded = True
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Set the Details group Text
Public Sub SetDetailsText(sDetails As String)
    On Error Resume Next
    m_DetailsGroupText = sDetails
    m_DetailsGroup.bExpanded = True
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Hide the Details Group
Public Sub HideDetailsGroup()
    On Error Resume Next
    m_bDetailsGroup = False
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Opens a url link
Public Function OpenLink(sLink As String) As Long
    On Error Resume Next
    OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
    Err.Clear
End Function
' Desc: try to explain where the hell does all this come from.
'Public Sub About()
'    on error resume next
'    MsgBox "isExplorerBar Control." & vbCrLf & "Developed By: Fred.cpp" & vbCrLf & "HomePage: http://mx.geocities.com/fred_cpp/isexplorerar.htm" & vbCrLf & "Description: this is a control that emulates almost all the functionality of the standard " & vbCrLf & "Windows Explorer Bar. Uses the Windows Theme currently installed.", vbInformation, "isExplorerBar"
'    Err.Clear
'End Sub
' Desc: Clear all the structure of the Control
Public Sub ClearStructure()
    On Error Resume Next
    'Clear all the icons and groups
    Dim ni As Integer
    Dim tmpCtl As Variant
    Dim btmpAllowUpdates As Boolean
    'Clear Special Group Items
    m_bSpecialGroup = False
    ReDim m_SpecialGroup.items(0)
    m_SpecialGroup.iItemsCount = 0
    'Clear Details Group
    m_bDetailsGroup = False
    ReDim cGroups(0)
    'Clear groups
    'clear Childs
    For ni = m_pChild.LBound To m_pChild.UBound
        If ni <> 0 Then
            For Each tmpCtl In UserControl.ContainedControls
                If tmpCtl.Name = m_pChild(ni).Tag Then
                    tmpCtl.Visible = False
                End If
            Next
        End If
        Unload m_pChild(ni)
    Next
    'Clear Groups
    'Clear Counter
    iGroups = 0
    'Refresh Control
    btmpAllowUpdates = m_AllowRedraw
    UserControl.MousePointer = 0
    m_AllowRedraw = True
    UserControl_Paint
    UserControl.Refresh
    m_AllowRedraw = btmpAllowUpdates
    Err.Clear
End Sub
' Desc: Clear all the structure of the Selected group
'       if you will change lots of groups, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub ClearGroup(Group)
    On Error Resume Next
    Dim iGroupIndex As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    'Clear all the icons in the selected group
    If iGroupIndex = -1 Then
        'clear special group Items
        ReDim m_SpecialGroup.items(0)
    Else
        'Clear a normal group
        ReDim cGroups(iGroupIndex).items(0)
        cGroups(iGroupIndex).iItemsCount = 0
        ' added by anele mbanga
        If TypeName(cGroups(iGroupIndex).pChild) = "PictureBox" Then
            cGroups(iGroupIndex).pChild.Visible = False
            Set cGroups(iGroupIndex).pChild = Nothing
            cGroups(iGroupIndex).ChildHidden = True
        ElseIf TypeName(cGroups(iGroupIndex).pChild) = "Frame" Then
            cGroups(iGroupIndex).pChild.Visible = False
            Set cGroups(iGroupIndex).pChild = Nothing
            cGroups(iGroupIndex).ChildHidden = True
        End If
    End If
    UserControl_Paint
    Err.Clear
End Sub
Public Sub HideGroup(sGroup)
    On Error Resume Next
    Dim iGroupIndex As Integer
    iGroupIndex = GetGroupsByKeyN(sGroup)
    'Clear all the icons in the selected group
    If iGroupIndex = -1 Then
        'clear special group Items
        ReDim m_SpecialGroup.items(0)
    Else
        'Clear a normal group
        ReDim cGroups(iGroupIndex).items(0)
        cGroups(iGroupIndex).iItemsCount = 0
        cGroups(iGroupIndex).Key = ""
        cGroups(iGroupIndex).Caption = ""
        ' added by anele mbanga
        If TypeName(cGroups(iGroupIndex).pChild) = "PictureBox" Then
            cGroups(iGroupIndex).pChild.Visible = False
            Unload m_pChild(iGroupIndex)
        ElseIf TypeName(cGroups(iGroupIndex).pChild) = "Frame" Then
            cGroups(iGroupIndex).pChild.Visible = False
            Unload m_pChild(iGroupIndex)
        End If
    End If
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Clear all the structure of the Selected group
'       if you will change lots of groups, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetGroupChild(Group, pChild As PictureBox, Optional pChildPointer As Integer = 1)
    On Error Resume Next
    Dim iGroupIndex As Integer
    Dim tmpCtl As Variant
    pChild.BorderStyle = 0
    iGroupIndex = GetGroupsByKeyN(Group)
    'Setup the Item Child.
    If iGroupIndex = -1 Then
        Set m_SpecialGroup.pChild = pChild 'ReDim m_SpecialGroup.items(0)
        pChild.ScaleMode = 3
        pChild.MousePointer = pChildPointer    'set Pointer
    Else
        'Clear a normal group
        'ReDim cGroups(iGroupIndex).items(0)
        If TypeName(cGroups(iGroupIndex).pChild) = "PictureBox" Then
            cGroups(iGroupIndex).pChild.Visible = True
            cGroups(iGroupIndex).pChild.Tag = cGroups(iGroupIndex).pChild.Name
            cGroups(iGroupIndex).pChild.ZOrder 0
        ElseIf TypeName(cGroups(iGroupIndex).pChild) = "Frame" Then
            cGroups(iGroupIndex).pChild.Visible = True
            cGroups(iGroupIndex).pChild.Tag = cGroups(iGroupIndex).pChild.Name
            cGroups(iGroupIndex).pChild.ZOrder 0
        Else
            Set cGroups(iGroupIndex).pChild = pChild
            pChild.ScaleMode = 3
            pChild.MousePointer = pChildPointer     'set Pointer
            pChild.ZOrder 0
            If cGroups(iGroupIndex).ChildHidden = True Then
                m_pChild(iGroupIndex).Visible = True
            Else
                Load m_pChild(iGroupIndex)
                Set m_pChild(iGroupIndex) = cGroups(iGroupIndex).pChild
                m_pChild(iGroupIndex).ScaleMode = 3
                m_pChild(iGroupIndex).Tag = pChild.Name
                m_pChild(iGroupIndex).AutoRedraw = True
                m_pChild(iGroupIndex).ZOrder 0
            End If
            'm_pChild(iGroupIndex).BackColor = cGroups(iGroupIndex).pChild.BackColor
        End If
    End If
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Expand an Especified group
Public Sub ExpandGroup(Group, Optional bExpand As Boolean = True)
    On Error Resume Next
    Dim iGroupIndex As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    'Colapse the selected group
    If iGroupIndex = -1 Then
        'Colapse Special Group
        If IsMissing(bExpand) Then bExpand = Not m_SpecialGroup.bExpanded
        m_SpecialGroup.bExpanded = bExpand
    ElseIf iGroupIndex = -2 Then
        'Colapse the selected Group
        If IsMissing(bExpand) Then bExpand = Not m_DetailsGroup.bExpanded
        m_DetailsGroup.bExpanded = bExpand
    Else
        'Colapse the selected Group
        If IsMissing("bExpand") Then bExpand = Not cGroups(iGroupIndex).bExpanded
        cGroups(iGroupIndex).bExpanded = bExpand
    End If
    UserControl_Paint
    Err.Clear
End Sub
'Desc:  This routine will change the text of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetGroupCaption(Group, sNewCaption As String, Optional sNewToolTip As String = "")
    On Error Resume Next
    'Set the icon of a item
    Dim iGroupIndex As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    If iGroupIndex = -3 Then
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -2 Then
        m_DetailsGroup.Caption = sNewCaption
        UserControl_Paint
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -1 Then
        m_SpecialGroup.Caption = sNewCaption
        m_SpecialGroup.ToolTip = sNewToolTip
        UserControl_Paint
        Err.Clear
        Exit Sub
    Else
        cGroups(iGroupIndex).Caption = sNewCaption
        cGroups(iGroupIndex).ToolTip = sNewToolTip
        UserControl_Paint
        Err.Clear
        Exit Sub
    End If
    Err.Clear
    Exit Sub
    'Item not found
    RaiseError "The group Doesn't Exist"
    Err.Clear
End Sub
'Desc:  This routine will change the icon of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetItemIcon(Group, Item, iNewIcon As Variant, Optional bUpdate As Boolean = True)
    On Error Resume Next
    'Set the icon of a item
    Dim iGroupIndex As Integer
    Dim iItemIndex As Integer
    Dim nj As Integer
    iNewIcon = GetIconIndex(iNewIcon)
    iGroupIndex = GetGroupsByKeyN(Group)
    If iGroupIndex = -3 Then
        RaiseError "The Group '" & Group & "' doesn't exist"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -2 Then
        RaiseError "Details Group hasn't Child Items!"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -1 Then
        iItemIndex = GetItemIndex(m_SpecialGroup, Item)
        If iItemIndex >= 1 Then
            m_SpecialGroup.items(iItemIndex).Icon = iNewIcon
            'RedrawItem iGroupIndex, iItemIndex, 1
            UserControl_Paint
            Err.Clear
            Exit Sub
        End If
    Else
        If GroupExists(iGroupIndex) Then
            iItemIndex = GetItemIndex(cGroups(iGroupIndex), Item)
            If iItemIndex >= 1 Then
                'We got the groupindex id and item index
                cGroups(iGroupIndex).items(iItemIndex).Icon = iNewIcon
                'RedrawItem iGroupIndex, iItemIndex, 1
                UserControl_Paint
                Err.Clear
                Exit Sub
            End If
        Else
            RaiseError "The Group '" & Group & "' doesn't exist"
            Err.Clear
            Exit Sub
        End If
    End If
    'When we get here, there shure was an error shown in func GetItemIndex
    'So we need not to raise another error here
    Err.Clear
End Sub
'Desc:  This routine will change the text of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetItemText(Group, Item, sNewCaption As String, Optional sNewToolTip As String = "")
    On Error Resume Next
    'Set the text of a item
    Dim iGroupIndex As Integer
    Dim iItemIndex As Integer
    Dim nj As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    If iGroupIndex = -3 Then
        'RaiseError "The Group '" & Group & "' doesn't exist"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -2 Then
        RaiseError "Details Group hasn't Child Items!"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -1 Then
        iItemIndex = GetItemIndex(m_SpecialGroup, Item)
        If iItemIndex >= 1 Then
            m_SpecialGroup.items(iItemIndex).Caption = sNewCaption
            m_SpecialGroup.items(iItemIndex).ToolTip = sNewToolTip
            'RedrawItem iGroupIndex, iItemIndex, 1
            UserControl_Paint
            Err.Clear
            Exit Sub
        End If
    Else
        If GroupExists(iGroupIndex) Then
            iItemIndex = GetItemIndex(cGroups(iGroupIndex), Item)
            If iItemIndex >= 1 Then
                'We got the groupindex id and item index
                cGroups(iGroupIndex).items(iItemIndex).Caption = sNewCaption
                cGroups(iGroupIndex).items(iItemIndex).ToolTip = sNewToolTip
                'RedrawItem iGroupIndex, iItemIndex, 1
                UserControl_Paint
                Err.Clear
                Exit Sub
            End If
        Else
            'RaiseError "The Group '" & Group & "' doesn't exist"
            Err.Clear
            Exit Sub
        End If
    End If
    'When we get here, there shure was an error shown in func GetItemIndex
    'So we need not to raise another error here
    Err.Clear
End Sub
'Desc:  this function disables drawing in the control.
'       Useful if you will change the entire structure
'       and don't want to slow down the execution with
'       multiple redraws.
'       Example:
'       isExplorerBar1.DisableUdates
'       for i = 1 to List1.listcount
'           isExplorerBar1.additem "MyGroupName","Action" & i, list1.list(i)
'       next i
'       isExplorerBar1.DisableUdates False
Public Sub DisableUpdates(Optional bDisable As Boolean = True)
    On Error Resume Next
    'Set the internal Variable
    m_AllowRedraw = Not bDisable
    'If the control has changed, I't a good Idea update the contents
    UserControl_Paint
    Err.Clear
End Sub
' Description: This Sub changes the Image shown in
'       the Details group.
'       To delete the previous Image, call the routine
'       without the detailsImage Parameter.
Public Sub SetDetailsImage(Optional ByVal detailsImage As Picture)
    On Error Resume Next
    Dim lmsize As Long
    '    lmsize = m_Width - 32
    Set m_DetailsPicture = detailsImage
    UserControl_Paint
    Err.Clear
End Sub
' Desc: Maybe you need check the Version while running
Public Function GetControlVersion() As String
    On Error Resume Next
    GetControlVersion = strCurrentVersion
    Err.Clear
End Function
' Desc: As Requested, Font Property
Public Property Set Font(newFont As StdFont)
    On Error Resume Next
    UserControl.Font.Name = newFont.Name
    UserControl.Font.Size = newFont.Size
    UserControl.Font.Charset = newFont.Charset
    Err.Clear
End Property
Public Property Get GroupsCount() As Long
    On Error Resume Next
    GroupsCount = iGroups
    Err.Clear
End Property
Public Property Get Font() As StdFont
    On Error Resume Next
    Set Font = UserControl.Font
    Err.Clear
End Property
Public Property Get UseUxThemeText() As Boolean
    On Error Resume Next
    UseUxThemeText = UxThemeText
    Err.Clear
End Property
Public Property Let UseUxThemeText(bNewUseUxThemeText As Boolean)
    On Error Resume Next
    UxThemeText = bNewUseUxThemeText
    UserControl_Paint
    Err.Clear
End Property
Public Function GetGroupCaption(Group As Variant) As String
    On Error Resume Next
    'Set the icon of a item
    Dim iGroupIndex As Integer
    Dim sNewCaption As String
    iGroupIndex = GetGroupsByKeyN(Group)
    If iGroupIndex = -3 Then
        sNewCaption = ""
    ElseIf iGroupIndex = -2 Then
        sNewCaption = m_DetailsGroup.Caption
    ElseIf iGroupIndex = -1 Then
        sNewCaption = m_SpecialGroup.Caption
    Else
        sNewCaption = cGroups(iGroupIndex).Caption
    End If
    GetGroupCaption = sNewCaption
    Err.Clear
End Function
Public Sub SetItemKey(Group, Item, sNewKey As String, Optional sNewCaption As String, Optional sNewToolTip As String = "")
    On Error Resume Next
    'Set the text of a item
    Dim iGroupIndex As Integer
    Dim iItemIndex As Integer
    Dim nj As Integer
    iGroupIndex = GetGroupsByKeyN(Group)
    If iGroupIndex = -3 Then
        RaiseError "The Group '" & Group & "' doesn't exist"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -2 Then
        RaiseError "Details Group hasn't Child Items!"
        Err.Clear
        Exit Sub
    ElseIf iGroupIndex = -1 Then
        iItemIndex = GetItemIndex(m_SpecialGroup, Item)
        If iItemIndex >= 1 Then
            m_SpecialGroup.items(iItemIndex).Key = sNewKey
            If Len(sNewCaption) > 0 Then m_SpecialGroup.items(iItemIndex).Caption = sNewCaption
            If Len(sNewToolTip) > 0 Then m_SpecialGroup.items(iItemIndex).ToolTip = sNewToolTip
            UserControl_Paint
            Err.Clear
            Exit Sub
        End If
    Else
        If GroupExists(iGroupIndex) Then
            iItemIndex = GetItemIndex(cGroups(iGroupIndex), Item)
            If iItemIndex >= 1 Then
                'We got the groupindex id and item index
                cGroups(iGroupIndex).items(iItemIndex).Key = sNewKey
                If Len(sNewCaption) > 0 Then cGroups(iGroupIndex).items(iItemIndex).Caption = sNewCaption
                If Len(sNewCaption) > 0 Then cGroups(iGroupIndex).items(iItemIndex).ToolTip = sNewToolTip
                UserControl_Paint
                Err.Clear
                Exit Sub
            End If
        Else
            RaiseError "The Group '" & Group & "' doesn't exist"
            Err.Clear
            Exit Sub
        End If
    End If
    'When we get here, there shure was an error shown in func GetItemIndex
    'So we need not to raise another error here
    Err.Clear
End Sub
Private Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = "|", Optional TrimItems As Boolean = False) As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = -1
    If Len(StringMv) = 0 Then
        MvSearch = -1
        Err.Clear
        Exit Function
    End If
    TheFields = Split(StringMv, Delim)
    If TrimItems = True Then ArrayTrimItems TheFields
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Err.Clear
End Function
Private Sub ArrayTrimItems(varArray() As String)
    On Error Resume Next
    Dim uArray As Long
    Dim cArray As Long
    Dim lArray As Long
    uArray = UBound(varArray)
    lArray = LBound(varArray)
    For cArray = lArray To uArray
        varArray(cArray) = Trim$(varArray(cArray))
    Next
    Err.Clear
End Sub
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ' return the position of an item in an array by looping
    Dim ArrayTot As Long
    Dim ArrayStart As Long
    Dim arrayCnt As Long
    Dim strCur As String
    ArrayStart = LBound(varArray)
    ArrayTot = UBound(varArray)
    StrSearch = LCase$(Trim$(StrSearch))
    ArraySearch = -1
    For arrayCnt = ArrayStart To ArrayTot
        strCur = LCase$(varArray(arrayCnt))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
    Err.Clear
End Function
