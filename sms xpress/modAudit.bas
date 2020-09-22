Attribute VB_Name = "modAudit"
Option Explicit
Public Enum ExcelFileFormat
    CSV = 1
    DBF4 = 2
    Html = 3
    TextMSDOS = 4
    TextPrinter = 5
    TextWindows = 6
    XMLSpreadsheet = 7
End Enum
Private ViewHeadings() As String
Private Const LB_DELETESTRING = &H182
Private Const CB_DELETESTRING = &H144
Private Const HWND_TOPMOST = -1
Private Const TV_FIRST As Long = &H1100
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Private Const TVGN_ROOT As Long = &H0
Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH As Long = 260&
Private Const SWP_NOMOVE As Long = 2
Private Const SWP_NOSIZE As Long = 1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
Private Const LB_ADDSTRING = &H180
Private Const CB_ADDSTRING = &H143
Private Const flags As Long = SWP_NOMOVE Or SWP_NOSIZE
Public activeMonth As String
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
End Type
Public Enum FileOps
    foDelete = 0
    foMove = 1
    foCopy = 2
    foRename = 3
End Enum
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const WM_USER = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)
Private Const FO_DELETE = &H3
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_RENAME As Long = &H4
Private Const FOF_ALLOWUNDO = &H40
Private blnAboveVer4 As Boolean
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const CB_FINDSTRINGEXACT = &H158
Private Const LVM_SETITEMCOUNT As Long = 4096 + 47
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Declare Function apiGetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetVersionExA Lib "KERNEL32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function FindWindowByTitle Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SetLocaleInfo Lib "KERNEL32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long
Private Declare Function GetComputerName Lib "Kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SQLAllocHandle Lib "odbc32.dll" (ByVal HandleType As Integer, ByVal InputHandle As Long, OutputHandlePtr As Long) As Integer
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal EnvironmentHandle As Long, ByVal Direction As Integer, ByVal ServerName As String, ByVal BufferLength1 As Integer, NameLength1Ptr As Integer, ByVal Description As String, ByVal BufferLength2 As Integer, NameLength2Ptr As Integer) As Integer
Private Declare Function SQLFreeHandle Lib "odbc32.dll" (ByVal HandleType As Integer, ByVal Handle As Long) As Integer
Private Declare Function SQLSetEnvAttr Lib "odbc32.dll" (ByVal EnvironmentHandle As Long, ByVal EnvAttribute As Long, ByVal ValuePtr As Long, ByVal StringLength As Long) As Integer
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Const ODBC_ADD_SYS_DSN      As Long = 4
Private Const SQL_NULL_HANDLE       As Long = 0
Private Const SQL_HANDLE_ENV        As Long = 1
Private Const SQL_FETCH_NEXT        As Long = 1
Private Const SQL_FETCH_FIRST       As Long = 2
Private Const SQL_SUCCESS           As Long = 0
Private Const SQL_ATTR_ODBC_VERSION As Long = 200
Private Const SQL_OV_ODBC2          As Long = 2
Private Const SQL_IS_INTEGER        As Long = -6
Private nRetCode  As Long
Private mServerName As String
Private mUserName As String
Private mPort As Long
Private mPassWord As String
Private mDatabaseName As String
Private mIsConnectionOpen As Boolean
Private mIsDatabaseOpen As Boolean
Private mValuesToUpdate As String
Private mInteractionType As Integer
Private Const LINE_BREAK As String = "\n"  ' Try "<br>". String to replace line breaks in text fields
Public usrPassword As String
Public Enum FindWhere
    search_Text = 0
    search_SubItem = 1
    Search_Tag = 2
End Enum
Public Enum SearchType
    search_Partial = 1
    search_Whole = 0
End Enum
Public vbNext As Long
Public vbClose As Long
Public vbSnooze As Long
Public vbSearch As Long
Public vbTips As Long
Public vbOptions As Long
Public vbYesToAll As Long
Public DateError As Integer
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Running As Boolean
Public resp As Integer
Public myModules As String
Private RetAnswer As Long
Public Intel As New clsIntellisense
Public myFields As String
Public myValues As String
Public UsrName As String
Public strPath As String
Public usrTelephone As String
Public usrEmail As String
Public usrFullName As String
Public usrOldPwd As String
Public usrExpiry As String
Public usrOnline As String
Public usrComputer As String
Public qrySql As String
Public LedgerHeading As String
Public ActiveSystem As String
Public IsGssc As String
Public IsNew As Boolean
Public myDescription As String
Public myNote As String
Public IsLogged As Boolean
Public MemberOf As String
Public OldPwd As String
Public Expiry As String
Public xLine As New Collection
Public drFound As Long
Public canOverwrite As Boolean
Public updateCnt As Long
Public Enum Ebt_Type
    BasEbt = 0
    SapEbt = 1
    PerEbt = 2
End Enum
Public FM As String
Public VM As String
Public NL As String

Sub Main()
    On Error Resume Next
    InitCommonControls
    'If App.PrevInstance Then
    '    Call MyPrompt("eFas is currently running, you can access it from the task bar.", "o", "e", "AMS")
    '    End
    'End If
    If SetDate = False Then
        UnloadAllForms
    End If
    FM = Chr$(254)
    VM = Chr$(253)
    NL = Chr$(13) + Chr$(10)
    Load frmSMSExchange
    Pause 2
    Err.Clear
End Sub
Public Sub UnloadAllForms()
    On Error Resume Next
    Dim oForm As Form
    For Each oForm In Forms
        Unload oForm
        Set oForm = Nothing
    Next
    End
    Err.Clear
End Sub
Sub FormMoveToNextControl(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
    Case 27   ' escape key
        SendKeys "+{TAB}"
        KeyAscii = 0
        DoEvents
    Case vbKeyReturn          ' catch return key
        Select Case TypeName(Screen.ActiveControl)
        Case "CheckBox", "ComboBox", "MaskEdBox", "DTPicker"
            SendKeys "{TAB}"      ' send tab which changes the element on form
            KeyAscii = 0
            DoEvents
        Case "TextBox"
            If Screen.ActiveControl.MultiLine = False Then
                SendKeys "{TAB}"      ' send tab which changes the element on form
                KeyAscii = 0
                DoEvents
            End If
        End Select
    End Select
    Err.Clear
End Sub
Sub RightAlignThese(lstReport As ListView)
    On Error Resume Next
    lstReport.ColumnHeaders(Val(LstViewColumnPosition(lstReport, "payment no"))).Alignment = lvwColumnRight
    lstReport.ColumnHeaders(Val(LstViewColumnPosition(lstReport, "total allocations"))).Alignment = lvwColumnRight
    lstReport.ColumnHeaders(Val(LstViewColumnPosition(lstReport, "total invoices"))).Alignment = lvwColumnRight
    lstReport.FullRowSelect = True
    lstReport.Refresh
    Err.Clear
End Sub
Sub ResetFilter(lstReport As ListView, lstValues As ListView, cboBox As ComboBox, chkRemove As CheckBox)
    On Error Resume Next
    Dim strColmn As String
    strColmn = LstViewColNames(lstReport)
    LstBoxFromMV cboBox, strColmn, ","
    lstValues.ListItems.Clear
    chkRemove.Value = 0
    StatusMessage lstReport.Parent, lstReport.ListItems.Count & " record(s) selected."
    Err.Clear
End Sub
Public Function LstBoxFindExactItemAPI(lstBox As Variant, ByVal sSearch As String) As Long
    On Error Resume Next
    Select Case TypeName(lstBox)
    Case "ListBox"
        LstBoxFindExactItemAPI = SendMessage(lstBox.hWnd, LB_FINDSTRINGEXACT, 0&, ByVal sSearch$)
    Case "ComboBox"
        LstBoxFindExactItemAPI = SendMessage(lstBox.hWnd, CB_FINDSTRINGEXACT, 0&, ByVal sSearch$)
    End Select
    Err.Clear
End Function
Public Function SpellCheck(ByVal strText As String, Optional ByVal blnSupressMsg As Boolean = True) As String
    On Error Resume Next
    Dim sTmpString As String
    Dim Speller As Variant
    SpellCheck = strText
    If strText = "" Then
        If blnSupressMsg = False Then
            resp = MyPrompt("There is nothing to spell check.", "o", "w", App.ProductName)
        End If
    Err.Clear
        Exit Function
    End If
    Set Speller = CreateObject("Word.Basic")
    With Speller
        .appminimize
        .filenew
        .Insert strText
        .editselectall
        .ToolsSpelling
        .editselectall
        sTmpString = .Selection()
        .fileexit 2
        .Quit
    End With
    sTmpString = Left$(sTmpString, Len(sTmpString) - 1)
    sTmpString = Replace$(sTmpString, Chr$(13), NL)
    If sTmpString = "" Then
        SpellCheck = strText
    Else
        SpellCheck = sTmpString
    End If
    Err.Clear
End Function
Public Function TimeDiff(ByVal LesserTime As String, ByVal GreaterTime As String, Optional ByVal LesserTimeDate As String = "01/01/2001", Optional ByVal GreaterTimeDate As String = "01/01/2001") As String
    On Error Resume Next
    Dim Thrs As Single
    Dim Tmins As Single
    Dim Tsecs As Single
    Dim ReturnTime As String
    Dim DTcheck As Single
    Dim i As Integer
    Dim Var_tmp_Hour As Integer
    DTcheck = DateDiff("d", LesserTimeDate, GreaterTimeDate)
    Var_tmp_Hour = Hour(GreaterTime)
    If DTcheck > 0 Then
        For i = 1 To DTcheck
            Var_tmp_Hour = Var_tmp_Hour + 24
        Next
    ElseIf DTcheck < 0 Then
        TimeDiff = "00:00"
    Err.Clear
        Exit Function
    End If
    Thrs = Var_tmp_Hour - Hour(LesserTime)
    If Thrs > 0 Then
        Tmins = Minute(GreaterTime) - Minute(LesserTime)
        If Tmins < 0 Then
            Tmins = Tmins + 60
            Thrs = Thrs - 1
        End If
    ElseIf Thrs = 0 And Minute(GreaterTime) >= Minute(LesserTime) Then
        Tmins = Minute(GreaterTime) - Minute(LesserTime)
    Else
        TimeDiff = "00:00"
    Err.Clear
        Exit Function
    End If
    Tsecs = Second(GreaterTime) - Second(LesserTime)
    If Tsecs < 0 Then
        Tsecs = Tsecs + 60
        If Tmins > 0 Then Tmins = Tmins - 1 Else Tmins = 59
    End If
    TimeDiff = StrFormat(Thrs, "R%2") & ":" & StrFormat(Tmins, "R%2")
    '& ":" & Tsecs
    Err.Clear
End Function
Public Function StrFormat(ByVal strData As String, ByVal sFormat As String, Optional ByVal StrType As String = "") As String
    On Error Resume Next
    Dim filler As String
    Dim field_length As Integer
    Dim diff_length As Integer
    Dim Justif As String
    Dim ResultingFieldLength As Integer
    Justif = Left$(sFormat, 1)
    filler = Mid$(sFormat, 2, 1)
    ResultingFieldLength = CInt(Mid$(sFormat, 3))
    Select Case filler
    Case "#"
        filler = " "
    Case "%"
        filler = "0"
    End Select
    Select Case strData
    Case ""
        Select Case UCase$(StrType)
        Case "M"
            strData = "0.00"
        End Select
    End Select
    field_length = Len(strData)
    Select Case field_length
    Case Is >= ResultingFieldLength
        StrFormat = Left$(strData, ResultingFieldLength)
    Case Else
        diff_length = ResultingFieldLength - field_length
        filler = String$(diff_length, filler)
        Select Case UCase$(Justif)
        Case "R"
            StrFormat = Concat(filler, strData)
        Case Else
            StrFormat = Concat(strData, filler)
        End Select
    End Select
    Err.Clear
End Function
Sub LstViewCheckFromMv(LstView As ListView, ByVal lstViewPos As Long, ByVal StrUseMv As String, Optional ByVal Delim As String = "", Optional boolCheck As Boolean = True, Optional useColor As Long = vbBlack, Optional bShow As Boolean = False)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim lstPos As Long
    Dim useData() As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    ' uncheck all items at first
    lstTot = LstView.ListItems.Count
    For lstCnt = 1 To lstTot
        LstView.ListItems(lstCnt).Checked = Not boolCheck
    Next
    lstTot = StrParse(useData, StrUseMv, Delim)
    For lstCnt = 1 To lstTot
        Select Case lstViewPos
        Case 1
            lstPos = LstViewFindItem(LstView, useData(lstCnt), search_Text, search_Whole)
        Case Else
            lstPos = LstViewFindItem(LstView, useData(lstCnt), search_SubItem, search_Whole)
        End Select
        If lstPos > 0 Then
            LstView.ListItems(lstPos).Checked = boolCheck
            LstView.ListItems(lstPos).ForeColor = useColor
        End If
    Next
    Err.Clear
End Sub
Public Function RemoveAlpha(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = UCase$(StrValue)
    StrValue = Replace$(StrValue, "A", "")
    StrValue = Replace$(StrValue, "B", "")
    StrValue = Replace$(StrValue, "C", "")
    StrValue = Replace$(StrValue, "D", "")
    StrValue = Replace$(StrValue, "E", "")
    StrValue = Replace$(StrValue, "F", "")
    StrValue = Replace$(StrValue, "G", "")
    StrValue = Replace$(StrValue, "H", "")
    StrValue = Replace$(StrValue, "I", "")
    StrValue = Replace$(StrValue, "J", "")
    StrValue = Replace$(StrValue, "K", "")
    StrValue = Replace$(StrValue, "L", "")
    StrValue = Replace$(StrValue, "M", "")
    StrValue = Replace$(StrValue, "N", "")
    StrValue = Replace$(StrValue, "O", "")
    StrValue = Replace$(StrValue, "P", "")
    StrValue = Replace$(StrValue, "Q", "")
    StrValue = Replace$(StrValue, "R", "")
    StrValue = Replace$(StrValue, "S", "")
    StrValue = Replace$(StrValue, "T", "")
    StrValue = Replace$(StrValue, "U", "")
    StrValue = Replace$(StrValue, "V", "")
    StrValue = Replace$(StrValue, "W", "")
    StrValue = Replace$(StrValue, "X", "")
    StrValue = Replace$(StrValue, "Y", "")
    StrValue = Replace$(StrValue, "Z", "")
    RemoveAlpha = StrValue
    Err.Clear
End Function
Public Function GetSysDir() As String
    On Error Resume Next
    Dim lpBuffer As String * 255
    Dim Length As Long
    Length = apiGetSystemDirectory(lpBuffer, Len(lpBuffer))
    GetSysDir = Left$(lpBuffer, Length)
    Err.Clear
End Function
Public Sub TextBoxHiLite(TxtBox As Variant)
    On Error Resume Next
    With TxtBox
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
    Err.Clear
End Sub
Public Function School_CleanName(ByVal strSchool As String) As String
    On Error Resume Next
    strSchool = Replace$(strSchool, "Primary School", "")
    strSchool = Replace$(strSchool, "Secondary School", "")
    strSchool = Replace$(strSchool, "High School", "")
    strSchool = Replace$(strSchool, "Primary", "")
    strSchool = Replace$(strSchool, "Sekondêr", "")
    strSchool = Replace$(strSchool, "Secondary", "")
    strSchool = Replace$(strSchool, "School", "")
    strSchool = Replace$(strSchool, "Skool", "")
    strSchool = Replace$(strSchool, "High", "")
    strSchool = Replace$(strSchool, "Primêre", "")
    strSchool = Replace$(strSchool, "Primêr", "")
    strSchool = Replace$(strSchool, "Hoërskool", "")
    strSchool = Replace$(strSchool, "Hoër", "")
    strSchool = Replace$(strSchool, "Laerskool", "")
    strSchool = Replace$(strSchool, "Primer", "")
    strSchool = Replace$(strSchool, "Intermediary", "")
    strSchool = Replace$(strSchool, "  ", " ")
    School_CleanName = Trim$(strSchool)
    Err.Clear
End Function
Public Function SequenceFinder(ByVal InString As String, Optional ByVal Delimiter As String = ",") As String
    On Error Resume Next
    Dim workstring As String
    Dim nums() As String
    Dim i As Long
    Dim j As Long
    Dim StartNum As Long
    Dim EndNum As Long
    Dim n1 As Long
    Dim strRslt As String
    strRslt = ""
    InString = MvRemoveDuplicates(InString, Delimiter)
    nums = Split(InString, Delimiter)
    ShellSort nums
    StartNum = nums(0)
    n1 = StartNum
    j = UBound(nums)
    If j = 1 Then
        strRslt = nums(j)
    Else
        For i = 1 To j
            If nums(i) <> n1 + 1 Then
                ' is next number in sequence?
                EndNum = nums(i - 1)
                ' if not, then last number ended sequence
                strRslt = strRslt & PrintSequences(StartNum, EndNum) & vbNewLine
                ' update the number counter
                StartNum = nums(i)
                ' get ready for next sequence
                n1 = StartNum
                EndNum = n1
            Else
                n1 = nums(i)
                ' this is a sequence where n1 is the current number in the sequence
                EndNum = n1
            End If
        Next
        'EndNum = nums(j)
        strRslt = strRslt & PrintSequences(StartNum, EndNum)  ' last sequence
    End If
    SequenceFinder = strRslt
    Err.Clear
End Function
Private Function PrintSequences(StartNum As Long, EndNum As Long) As String
    On Error Resume Next
    If StartNum = EndNum Then
        ' if start number and end number the same, only print the number
        PrintSequences = CStr(StartNum)
        ' replace msgbox with whatever your print or display routine is
    ElseIf EndNum < StartNum Then
        PrintSequences = CStr(StartNum)
    Else
        PrintSequences = CStr(StartNum) & " - " & CStr(EndNum)  ' otherwise print the start and end number
    End If
    Err.Clear
End Function
Sub ShellSort(vArray As Variant)
    On Error Resume Next
    Dim lLoop1 As Long
    Dim lHold As Long
    Dim lHValue As Long
    Dim lTemp As Long
    lHValue = LBound(vArray)
    Do
        lHValue = 3 * lHValue + 1
    Loop Until lHValue > UBound(vArray)
    Do
        lHValue = lHValue / 3
        For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
            lTemp = Val(vArray(lLoop1))
            lHold = Val(lLoop1)
            Do While Val(vArray(lHold - lHValue)) > lTemp
                vArray(lHold) = vArray(lHold - lHValue)
                lHold = lHold - lHValue
                If lHold < lHValue Then Exit Do
            Loop
            vArray(lHold) = lTemp
        Next
    Loop Until lHValue = LBound(vArray)
    Err.Clear
End Sub
Function MVLastItem(ByVal StringMv As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spValues() As String
    Dim spTotal As Long
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(spValues, StringMv, Delim)
    spTotal = UBound(spValues)
    MVLastItem = spValues(spTotal)
    Err.Clear
End Function
Public Function Rest(ByVal strData As String, Optional ByVal startPos As Long = 1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spData() As String
    Dim spCnt As Long
    Dim intLoop As Long
    Dim strL As String
    Rest = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    If Len(strData) = 0 Then
    Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case startPos
    Case -1
        Rest = Trim$(spData(spCnt))
    Case Else
        strL = ""
        For intLoop = startPos To spCnt
            strL = spData(intLoop)
            Rest = StringsConcat(Rest, strL, Delim)
        Next
        Rest = RemoveDelim(Rest, Delim)
    End Select
    Err.Clear
End Function
Sub LstViewSumTime(lstReport As ListView, ParamArray vColumnNames())
    On Error Resume Next
    Dim colPos As Long
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim strSum As String
    Dim RowPos As Long
    Dim colName As String
    Dim vColumnName As Variant
    Dim colTot As Long
    For Each vColumnName In vColumnNames
        colName = CStr(vColumnName)
        colPos = LstViewColumnPosition(lstReport, colName)
        strSum = "00:00"
        rsTot = lstReport.ListItems.Count
        For rsCnt = 1 To rsTot
            spLine = LstViewGetRow(lstReport, rsCnt)
            If spLine(1) = "Totals" Then
            Else
                strSum = TimeAdd(strSum, spLine(colPos))
            End If
        Next
        colTot = UBound(spLine)
        RowPos = LstViewFindItem(lstReport, "Totals", search_Text, search_Whole)
        Select Case RowPos
        Case 0
            ReDim spLine(colTot)
            spLine(1) = "Totals"
            spLine(colPos) = strSum
            Call LstViewUpdate(spLine, lstReport, "")
        Case Else
            spLine = LstViewGetRow(lstReport, RowPos)
            spLine(colPos) = strSum
            RowPos = LstViewUpdate(spLine, lstReport, CStr(RowPos))
            lstReport.ListItems(RowPos).EnsureVisible
        End Select
    Next
    Err.Clear
End Sub
Public Function TimeAdd(ByVal tPrevious As String, ByVal tCurrent As String)
    On Error Resume Next
    Dim tpHH As String
    Dim tpMM As String
    Dim tcHH As String
    Dim tcMM As String
    Dim mmSum As String
    Dim hhSum As String
    Dim mmRem As String
    tpHH = MvField(tPrevious, 1, ":")
    tpMM = MvField(tPrevious, 2, ":")
    tcHH = MvField(tCurrent, 1, ":")
    tcMM = MvField(tCurrent, 2, ":")
    hhSum = Val(tpHH) + Val(tcHH)
    mmSum = Val(tpMM) + Val(tcMM)
    Select Case mmSum
    Case 60
        hhSum = Val(hhSum) + 1
        mmSum = "00:00"
    Case Is > 60
        hhSum = Val(hhSum) + 1
        mmSum = Val(mmSum) - 60
    End Select
    TimeAdd = StrFormat(hhSum, "R%2") & ":" & StrFormat(mmSum, "R%2")
    Err.Clear
End Function
Public Function MvSort_Numbers(ByVal StrString As String, Optional ByVal Delimiter As String = ",") As String
    On Error Resume Next
    Dim sortArray As Variant
    sortArray = Split(StrString, Delimiter)
    ShellSort sortArray
    MvSort_Numbers = MvFromArray(sortArray, Delimiter, 0)
    Err.Clear
End Function
Public Function RecycleFile(OwnerForm As Form, fromPaths As String, Optional toPaths As String = "", Optional intPerform As FileOps = foCopy) As Boolean
    On Error Resume Next
    DoEvents
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    With FileOperation
        .hWnd = OwnerForm.hWnd
        Select Case intPerform
        Case 0
            .wFunc = FO_DELETE
        Case 1
            .wFunc = FO_MOVE
        Case 2
            .wFunc = FO_COPY
        Case 3
            .wFunc = FO_RENAME
        End Select
        .pFrom = fromPaths & vbNullChar & vbNullChar
        '.fFlags = FOF_SIMPLEPROGRESS Or FOF_ALLOWUNDO Or FOF_CREATEPROGRESSDLG
        .fFlags = FOF_ALLOWUNDO
        If Len(toPaths) > 0 Then
            .pTo = toPaths & vbNullChar & vbNullChar
        End If
    End With
    lReturn = SHFileOperation(FileOperation)
    RecycleFile = True
    If lReturn <> 0 Then
        ' Operation failed
        RecycleFile = False
    Else
        If FileOperation.fAnyOperationsAborted <> 0 Then
            RecycleFile = False
        End If
    End If
    Err.Clear
End Function
Sub LstViewFilterLike(lstReport As ListView, ByVal ColumnName As String, ByVal ColumnValue As String, Optional Remove As Integer = 0)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim xCols As String
    Dim xPos As Long
    Dim spLine() As String
    Dim curValue As String
    xCols = LstViewColNames(lstReport)
    xPos = MvSearch(xCols, ColumnName, ",")
    If xPos = 0 Then Exit Sub
    ColumnValue = LCase$(ColumnValue)
    If ColumnValue = "(none)" Or ColumnValue = "(blank)" Then ColumnValue = ""
    ColumnValue = MvReplaceItem(ColumnValue, "(none)", "", VM)
    ColumnValue = MvReplaceItem(ColumnValue, "(blank)", "", VM)
    rsTot = lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLine = LstViewGetRow(lstReport, rsCnt)
        curValue = LCase$(Trim$(spLine(xPos)))
        If Remove = 0 Then
            If InStr(1, curValue, ColumnValue, vbTextCompare) = 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        Else
            If InStr(1, curValue, ColumnValue, vbTextCompare) > 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        End If
    Next
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Public Sub LstViewCopyChechecked(lstSource As ListView, lstTarget As ListView)
    On Error Resume Next
    Dim spLine() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim sHeads As String
    sHeads = LstViewColNames(lstSource)
    lstTarget.ListItems.Clear
    LstViewMakeHeadings lstTarget, sHeads
    spTot = lstSource.ListItems.Count
    LstViewSetMemory lstTarget, spTot
    For spCnt = 1 To spTot
        If lstSource.ListItems(spCnt).Checked = False Then GoTo NextLine
        spLine = LstViewGetRow(lstSource, spCnt)
        Call LstViewUpdate(spLine, lstTarget, "")
NextLine:
    Next
    ' align columns
    spTot = StrParse(spLine, sHeads, ",")
    For spCnt = 1 To spTot
        lstTarget.ColumnHeaders(spCnt).Alignment = lstSource.ColumnHeaders(spCnt).Alignment
    Next
    LstViewAutoResize lstTarget
    lstTarget.Refresh
    DoEvents
    Err.Clear
End Sub
Sub LstViewToWordTable(ByVal Header As String, ByVal Footer As String, ByVal strFileName As String, LstView As ListView, Optional ByVal xOrientation As String = "landscape")
    On Error Resume Next
    If LstView.ListItems.Count = 0 Then Exit Sub
    Dim objRange As Word.Range
    Dim objDocument As Word.Document
    Dim numColumns As Long
    Dim numColumn As Long
    Dim numRows As Long
    Dim lstData() As String
    Dim numRow As Long
    Dim colName As String
    Dim oWord As Word.Application
    Set oWord = New Word.Application
    If oWord Is Nothing Then
        Call intError("MS Word Error", "MS Word is not installed on your computer.")
    Err.Clear
        Exit Sub
    End If
    With oWord
        .ScreenUpdating = False
        .Options.CheckGrammarAsYouType = False
        .Options.CheckSpellingAsYouType = False
        .DisplayAlerts = wdAlertsNone
    End With
    numRows = LstView.ListItems.Count + 1
    numColumns = LstView.ColumnHeaders.Count
    colName = LstViewColNames(LstView)
    Call StrParse(lstData, colName, ",")
    ' create a new word document that will hold all details in topic file
    Set objDocument = oWord.Documents.Add
    With objDocument
        .Activate
        Select Case LCase$(Left$(xOrientation, 1))
        Case "l"
            .PageSetup.Orientation = wdOrientLandscape
        Case "p"
            .PageSetup.Orientation = wdOrientPortrait
        End Select
        .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = ProperCase(Header)
        .Sections(1).Headers(wdHeaderFooterPrimary).Range.Bold = True
        .Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Name = "Tahoma"
        .Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Size = 10
        .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = ProperCase(Footer)
        .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight
        .Sections(1).Footers(wdHeaderFooterPrimary).Range.Font.Name = "Tahoma"
        .Sections(1).Footers(wdHeaderFooterPrimary).Range.Font.Size = 8
        .PageSetup.TopMargin = CentimetersToPoints(2)
        .PageSetup.BottomMargin = CentimetersToPoints(2)
        .PageSetup.LeftMargin = CentimetersToPoints(2)
        .PageSetup.RightMargin = CentimetersToPoints(2)
        .PageSetup.Gutter = CentimetersToPoints(1)
        .PageSetup.HeaderDistance = CentimetersToPoints(1.25)
        .PageSetup.FooterDistance = CentimetersToPoints(1.25)
        .PageSetup.PageWidth = CentimetersToPoints(21)
        .PageSetup.PageHeight = CentimetersToPoints(29.7)
        .PageSetup.GutterPos = wdGutterPosLeft
    End With
    'objDocument.FitToPages
    Set objRange = objDocument.Range(Start:=0, End:=0)
    objDocument.Tables.Add Range:=objRange, numRows:=numRows, numColumns:=numColumns
    With objDocument.Tables(1)
        .Range.Font.Bold = False
        .Range.Font.Name = "Tahoma"
        .Range.Font.Size = 8
        .Rows(1).Range.Font.Bold = True
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Rows(1).HeadingFormat = True
    End With
    For numColumn = 1 To numColumns
        objDocument.Tables(1).cell(1, numColumn).Range.Text = lstData(numColumn)
        'If lstView.ColumnHeaders(numColumn).Alignment = lvwColumnRight Then
        '    objDocument.Tables(1).cell(1, numColumn).RightPadding = True
        'End If
    Next
    numRows = LstView.ListItems.Count
    For numRow = 1 To numRows
        lstData = LstViewGetRow(LstView, numRow)
        For numColumn = 1 To numColumns
            objDocument.Tables(1).cell(numRow + 1, numColumn).Range.Text = lstData(numColumn)
            'If lstView.ColumnHeaders(numColumn).Alignment = lvwColumnRight Then
            '    objDocument.Tables(1).cell(numRow + 1, numColumn).RightPadding = True
            'End If
        Next
    Next
    objDocument.Tables(1).Columns.AutoFit
    ' save document as a rich text document
    objDocument.SaveAs strFileName, wdFormatRTF
    objDocument.Close
    oWord.Quit
    Set oWord = Nothing
    Set objRange = Nothing
    Set objDocument = Nothing
ExitSub:
    Err.Clear
End Sub
Public Function intError(ByVal StrTitle As String, ByVal strmessage As String) As Integer
    On Error Resume Next
    intError = MyPrompt(strmessage, "o", "w", ProperCase(StrTitle))
    Err.Clear
End Function
Public Function MvDoubleQuote(ByVal strData As String, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    Dim sData() As String
    Dim tCnt As Integer
    Dim wCnt As Integer
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(sData, strData, Delim)
    wCnt = UBound(sData)
    For tCnt = 1 To wCnt
        sData(tCnt) = StringAdd(Quote, sData(tCnt), Quote)
    Next
    MvDoubleQuote = MvFromArray(sData, Delim)
    Err.Clear
End Function
Function LstViewToCSV(LstView As ListView, ByVal strFile As String) As String
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLine() As String
    Dim nPath As String
    Dim fColN As String
    Dim fRowN As String
    Dim fNumber As Long
    ' get the path of the file and create it if it does not exist
    strFile = FileName_Validate(strFile)
    nPath = FileToken(strFile, "p")
    If DirExists(nPath) = False Then
        MakeDirectory nPath
    End If
    If FileExists(strFile) = True Then Kill strFile
    DoEvents
    ' what are the column names
    fColN = LstViewColNames(LstView)
    fColN = MvDoubleQuote(fColN, ",")
    ' create temporal files from the records
    rsTot = LstView.ListItems.Count
    For rsCnt = 1 To rsTot
        spLine = LstViewGetRow(LstView, rsCnt)
        ' make a string of the row details
        fRowN = MvFromArray(spLine, FM)
        ' Quote each item of the string
        fRowN = MvDoubleQuote(fRowN, FM)
        ' ensure each item is separated by a comma
        fRowN = Replace$(fRowN, FM, ",")
        If FileExists(strFile) = False Then
            fNumber = FreeFile
            Open strFile For Output Access Write As #fNumber
                Print #fNumber, fColN
                Print #fNumber, fRowN
            Close #fNumber
        Else
            fNumber = FreeFile
            Open strFile For Append Access Write As #fNumber
                Print #fNumber, fRowN
            Close #fNumber
        End If
NextRecord:
    Next
    DoEvents
    LstViewToCSV = strFile
    Err.Clear
End Function
Function AuditTrailFile(OwnerForm As Form, ByVal strFile As String) As String
    On Error Resume Next
    Dim strLine As String
    Dim sRslt As String
    Dim sData As String
    Dim sLines() As String
    Dim rTot As Long
    Dim rCnt As Long
    Dim xStart As String
    Dim xEnd As String
    sData = FileData(strFile)
    If InStr(1, sData, "AUDIT TRAIL LIST", vbTextCompare) = 0 Then
        Call RecycleFile(OwnerForm, strFile, , foDelete)
        AuditTrailFile = ""
    Err.Clear
        Exit Function
    Else
        sRslt = ""
        sLines = Split(sData, vbNewLine)
        rTot = UBound(sLines)
        For rCnt = 0 To rTot
            strLine = Trim$(sLines(rCnt))
            If Len(strLine) = 0 Then GoTo NextLine
            If Left$(strLine, Len("1. TRANS NO FROM                   :")) = "1. TRANS NO FROM                   :" Then
                xStart = Trim$(Mid$(strLine, Len("1. TRANS NO FROM                   :") + 1))
            End If
            If Left$(strLine, Len("2. TRANS NO TO                     :")) = "2. TRANS NO TO                     :" Then
                xEnd = Trim$(Mid$(strLine, Len("2. TRANS NO TO                     :") + 1))
                Exit For
            End If
NextLine:
        Next
        If xStart = xEnd Then
            sRslt = xStart
        Else
            sRslt = xStart & " - " & xEnd
        End If
        AuditTrailFile = FileToken(strFile, "p") & "\AT " & sRslt & ".txt"
    End If
    Err.Clear
End Function
Public Function MvSum(ByVal StringMv As String, Optional ByVal Delim As String = "", Optional Moneytary As Boolean = False) As String
    On Error Resume Next
    Dim rslt As Double
    Dim lngCnt As Long
    Dim MV() As String
    Dim wCnt As Long
    Dim sStr As String
    rslt = 0
    If Len(Delim) = 0 Then Delim = VM
    Call StrParse(MV, StringMv, Delim)
    wCnt = UBound(MV)
    For lngCnt = 1 To wCnt
        sStr = ProperAmount(MV(lngCnt))
        rslt = rslt + CDbl(sStr)
    Next
    If Moneytary = True Then
        MvSum = MakeMoney(CStr(rslt))
    Else
        MvSum = ProperAmount(CStr(rslt))
    End If
    Err.Clear
End Function
Function ExtractMatches(ByVal strSchool As String, ByVal StrAllocation As String, Optional SchoolStartsWithDistrict As Boolean = True, Optional ShowAll As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spResp() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim xPos(100) As String
    Dim spWords() As String
    Dim spWTot As Long
    Dim spWCnt As Long
    Dim spWPer As Long
    Dim xRslt As Long
    If SchoolStartsWithDistrict = True Then strSchool = Trim$(Mid$(strSchool, 3))
    strSchool = Replace$(strSchool, "Primary School", "Ps")
    strSchool = Replace$(strSchool, "Secondary School", "Ss")
    strSchool = Replace$(strSchool, "High School", "Hs")
    strSchool = Replace$(strSchool, "Primary", "")
    strSchool = Replace$(strSchool, "Sekondêr", "")
    strSchool = Replace$(strSchool, "Secondary", "")
    strSchool = Replace$(strSchool, "School", "")
    strSchool = Replace$(strSchool, "Skool", "")
    strSchool = Replace$(strSchool, "High", "")
    strSchool = Replace$(strSchool, "Primêre", "")
    strSchool = Replace$(strSchool, "Primêr", "")
    strSchool = Replace$(strSchool, "Hoërskool", "")
    strSchool = Replace$(strSchool, "Hoër", "")
    strSchool = Replace$(strSchool, "Laerskool", "")
    strSchool = Replace$(strSchool, "Primer", "")
    strSchool = Replace$(strSchool, "Intermediary", "")
    strSchool = Replace$(strSchool, "  ", " ")
    strSchool = Trim$(strSchool)
    spWTot = StrParse(spWords, strSchool, " ")
    spTot = StrParse(spResp, StrAllocation, VM)
    ' loop through each allocation and find the strings
    For spCnt = 1 To spTot
        spStr = spResp(spCnt)
        spWPer = 0
        For spWCnt = 1 To spWTot
            If MvSearch(spStr, spWords(spWCnt), " ") > 0 Then
                spWPer = spWPer + 1
            End If
        Next
        If spWPer <> 0 Then
            xRslt = (spWPer / spWTot) * 100
            If xRslt >= 50 Then
                xPos(xRslt) = spStr
            End If
        End If
    Next
    spStr = ""
    For rsCnt = 100 To 50 Step -1
        If Len(xPos(rsCnt)) > 0 Then
            spStr = spStr & xPos(rsCnt) & VM
        End If
    Next
    spStr = RemoveDelim(spStr, VM)
    If ShowAll = False Then
        ExtractMatches = MvField(spStr, 1, VM)
    Else
        ExtractMatches = spStr
    End If
    Err.Clear
End Function
Public Function AtPercentage(ByVal lngSmall As Long, ByVal lngBig As Long) As String
    On Error Resume Next
    Dim strX As String
    strX = (lngSmall / lngBig) * 100
    strX = ProperAmount(strX)
    strX = Fix(Val(strX))
    AtPercentage = Concat(strX, "%")
    Err.Clear
End Function
Function DistrictFullName(ByVal StrValue As String) As String
    On Error Resume Next
    DistrictFullName = ReadRecordToMv("districts", "ID", StrValue, "DistrictCode,Name", " - ")
    If Len(DistrictFullName) = 0 Then DistrictFullName = StrValue
    Err.Clear
End Function
Public Function YyyymmddToNormal(ByVal StrValue As String) As String
    On Error Resume Next
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    yyyy = Left$(StrValue, 4)
    mm = Mid$(StrValue, 5, 2)
    dd = Right$(StrValue, 2)
    YyyymmddToNormal = dd & "/" & mm & "/" & yyyy
    Err.Clear
End Function
Public Function SpellNumber(ByVal Mynumber As String) As String
    On Error Resume Next
    Dim Dollars As String
    Dim Cents As String
    Dim temp As String
    Dim DecimalPlace As Integer
    Dim Count As Integer
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' String representation of amount.
    Mynumber = Trim$(Mynumber)
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(Mynumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left$(Mid$(Mynumber, DecimalPlace + 1) & "00", 2))
        Mynumber = Trim$(Left$(Mynumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While Mynumber <> ""
        temp = GetHundreds(Right$(Mynumber, 3))
        If temp <> "" Then
            Dollars = temp & Place(Count) & Dollars
        End If
        If Len(Mynumber) > 3 Then
            Mynumber = Left$(Mynumber, Len(Mynumber) - 3)
        Else
            Mynumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Dollars
    Case ""
        Dollars = ""        '"No Dollars"
    Case "One"
        Dollars = "One"     '"One Dollar"
    Case Else
        Dollars = Dollars   ' & " Dollars"
    End Select
    Select Case Cents
    Case ""
        Cents = ""   '" and No Cents"
    Case "One"
        Cents = " One"  '" and One Cent"
    Case Else
        Cents = " " & Cents '  " and " & Cents & " Cents"
    End Select
    SpellNumber = Dollars & Cents
    Err.Clear
End Function
' Converts a number from 100-999 into text
Private Function GetHundreds(ByVal Mynumber As String) As String
    On Error Resume Next
    Dim Result As String
    Dim resTmp As String
    If Val(Mynumber) = 0 Then
    Err.Clear
        Exit Function
    End If
    Mynumber = Right$("000" & Mynumber, 3)
    ' Convert the hundreds place.
    If Mid$(Mynumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid$(Mynumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid$(Mynumber, 2, 1) <> "0" Then
        resTmp = GetTens(Mid$(Mynumber, 2))
        Select Case resTmp
        Case ""
            Result = Result & resTmp
        Case Else
            If Result <> "" Then
                Result = Result & "and " & resTmp
            Else
                Result = Result & resTmp
            End If
        End Select
        'Result = Result & GetTens(mid$(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid$(Mynumber, 3))
    End If
    GetHundreds = Result
    Err.Clear
End Function
' Converts a number from 10 to 99 into text.
Private Function GetTens(ByVal Tenstext As String) As String
    On Error Resume Next
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left$(Tenstext, 1)) = 1 Then
        ' If value between 10-19...
        Select Case Val(Tenstext)
        Case 10: Result = "Ten"
        Case 11: Result = "Eleven"
        Case 12: Result = "Twelve"
        Case 13: Result = "Thirteen"
        Case 14: Result = "Fourteen"
        Case 15: Result = "Fifteen"
        Case 16: Result = "Sixteen"
        Case 17: Result = "Seventeen"
        Case 18: Result = "Eighteen"
        Case 19: Result = "Nineteen"
        Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left$(Tenstext, 1))
        Case 2: Result = "Twenty "
        Case 3: Result = "Thirty "
        Case 4: Result = "Forty "
        Case 5: Result = "Fifty "
        Case 6: Result = "Sixty "
        Case 7: Result = "Seventy "
        Case 8: Result = "Eighty "
        Case 9: Result = "Ninety "
        Case Else
        End Select
        Result = Result & GetDigit(Right$(Tenstext, 1))        ' Retrieve ones place.
    End If
    GetTens = Result
    Err.Clear
End Function
' Converts a number from 1 to 9 into text.
Private Function GetDigit(ByVal Digit As String) As String
    On Error Resume Next
    Select Case Val(Digit)
    Case 1: GetDigit = "One"
    Case 2: GetDigit = "Two"
    Case 3: GetDigit = "Three"
    Case 4: GetDigit = "Four"
    Case 5: GetDigit = "Five"
    Case 6: GetDigit = "Six"
    Case 7: GetDigit = "Seven"
    Case 8: GetDigit = "Eight"
    Case 9: GetDigit = "Nine"
    Case Else: GetDigit = ""
    End Select
    Err.Clear
End Function
Public Function WeekEndsBetweenTwoDates(ByVal sDate As String, ByVal eDate As String) As Long
    On Error Resume Next
    Dim lngsDate As Long
    Dim lngeDate As Long
    Dim rsCnt As Long
    Dim outDate As String
    Dim weekEnds As Long
    Dim myDate As String
    weekEnds = 0
    lngsDate = Val(DateIconv(sDate))
    lngeDate = Val(DateIconv(eDate))
    For rsCnt = lngsDate To lngeDate
        myDate = DateOconv(CStr(rsCnt))
        myDate = Format$(CDate(myDate), "ddd mm yyyy")
        Select Case MvField(myDate, 1, " ")
        Case "Sat", "Sun"
            weekEnds = weekEnds + 1
        End Select
    Next
    WeekEndsBetweenTwoDates = weekEnds
    Err.Clear
End Function
Public Sub LstBoxUpdateAPI(lstBox As Variant, ParamArray items())
    On Error Resume Next
    Dim Item As Variant
    For Each Item In items
        If LstBoxFindExactItemAPI(lstBox, CStr(Item)) = -1 Then
            LstBoxAddItemAPI lstBox, CStr(Item)
        End If
    Next
    Set Item = Nothing
    Err.Clear
End Sub
Public Sub LstBoxAddItemAPI(lstBox As Variant, ParamArray cboItems())
    On Error Resume Next
    Dim cboItem As Variant
    Dim cboStr As String
    Select Case TypeName(lstBox)
    Case "ListBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            Call SendMessage(lstBox.hWnd, LB_ADDSTRING, 0&, ByVal cboStr$)
        Next
    Case "ComboBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            Call SendMessage(lstBox.hWnd, CB_ADDSTRING, 0&, ByVal cboStr$)
        Next
    End Select
    Err.Clear
End Sub
Public Function LstBoxToMV(lstBox As Variant, Optional ByVal Delim As String = "", Optional ByVal startPos As Integer = 0, Optional ByVal PutBrackets As Boolean = False) As String
    On Error Resume Next
    Dim lCnt As Long
    Dim sStr As String
    Dim xStr As String
    Dim totArray As Long
    sStr = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    If Val(startPos) <= 0 Then
        startPos = 1
    End If
    totArray = lstBox.ListCount - 1
    For lCnt = 0 To totArray
        xStr = lstBox.List(lCnt)
        xStr = Mid$(xStr, startPos)
        If PutBrackets = True Then
            xStr = StringsConcat("[", xStr, "]")
        End If
        Select Case lCnt
        Case totArray
            sStr = Concat(sStr, xStr)
        Case Else
            sStr = StringAdd(sStr, xStr, Delim)
        End Select
    Next
    LstBoxToMV = sStr
    Err.Clear
End Function
Sub LstViewRowFormat(LstView As ListView, RowPos As Long, Optional bChecked As Boolean = False, Optional ByVal bBold As Boolean = False, Optional ByVal cForeColor As Long = vbBlack, Optional rIcon As String = "", Optional ByVal tTooltip As String = "", Optional ByVal tTag As String = "")
    On Error Resume Next
    Dim numColumns As Long
    Dim numColumn As Long
    numColumns = LstView.ColumnHeaders.Count - 1
    LstView.ListItems(RowPos).Bold = bBold
    If Len(rIcon) > 0 Then
        LstView.ListItems(RowPos).Icon = rIcon
    End If
    LstView.ListItems(RowPos).Tag = tTag
    LstView.ListItems(RowPos).ToolTipText = tTooltip
    LstView.ListItems(RowPos).Checked = bChecked
    LstView.ListItems(RowPos).ForeColor = cForeColor
    For numColumn = 1 To numColumns
        LstView.ListItems(RowPos).ListSubItems(numColumn).Bold = bBold
        LstView.ListItems(RowPos).ListSubItems(numColumn).Tag = tTag
        LstView.ListItems(RowPos).ListSubItems(numColumn).ForeColor = cForeColor
        LstView.ListItems(RowPos).ListSubItems(numColumn).ToolTipText = tTooltip
    Next
    Err.Clear
End Sub
Public Function MvRemoveLastCharacter(ByVal StrString, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    rsTot = StrParse(spLine, StrString, Delim)
    For rsCnt = 1 To rsTot
        spLine(rsCnt) = Left$(spLine(rsCnt), Len(spLine(rsCnt)) - 1)
    Next
    MvRemoveLastCharacter = MvFromArray(spLine, Delim)
    Err.Clear
End Function
Function VerifyUserLogin(ByVal sUserName As String, ByVal sPassWord As String) As Boolean
    On Error Resume Next
    Dim sql As String
    sql = ReadRecordToMv("users", "UserName", sUserName, "UserPassword,Telephone,Email,FullName,Active,OldPasswords,Expiry,Modules,Online,Computer,Gssc", FM)
    If MvField(sql, 1, FM) = DecryptString(sPassWord) Then
        VerifyUserLogin = True
        usrPassword = MvField(sql, 1, FM)
        usrTelephone = MvField(sql, 2, FM)
        usrEmail = MvField(sql, 3, FM)
        usrFullName = MvField(sql, 4, FM)
        usrOldPwd = MvField(sql, 6, FM)
        usrExpiry = MvField(sql, 7, FM)
        If MvField(sql, 5, FM) = "0" Then VerifyUserLogin = False
        myModules = MvField(sql, 8, FM)
        'If Len(myModules) > 0 Then myModules = MvRemoveLastCharacter(myModules, ",")
        usrOnline = MvField(sql, 9, FM)
        usrComputer = MvField(sql, 10, FM)
        IsGssc = MvField(sql, 11, FM)
    Else
        VerifyUserLogin = False
        usrPassword = ""
        usrTelephone = ""
        usrEmail = ""
        usrFullName = ""
        usrOldPwd = ""
        usrExpiry = ""
        myModules = ""
        usrOnline = ""
        usrComputer = ""
        IsGssc = "0"
    End If
    Err.Clear
End Function
Public Function RN(Fromvariable As Variant, Optional ClearQuotes As Boolean = False)
    On Error Resume Next
    RN = Trim$(Fromvariable & "")
    If ClearQuotes = True Then
        RN = Trim$(Replace$(RN, Quote, ""))
    End If
    Err.Clear
End Function
Function MvRemoveDuplicates(ByVal StrMvString As String, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim spData() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim xCol As New Collection
    spData = Split(StrMvString, Delim)
    spTot = UBound(spData)
    For spCnt = 0 To spTot
        xCol.Add spData(spCnt), spData(spCnt)
    Next
    MvRemoveDuplicates = MvFromCollection(xCol, Delim)
    Err.Clear
End Function
Function MvFromCollection(objCollection As Collection, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    If Delim = "" Then Delim = VM
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        If xCnt = xTot Then
            sRet = sRet & objCollection.Item(xCnt)
        Else
            sRet = sRet & objCollection.Item(xCnt) & Delim
        End If
    Next
    MvFromCollection = sRet
    Err.Clear
End Function
Public Function DaoFldNames(ByVal dbRs As DAO.Recordset, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    fL = ""
    If Len(Delim) = 0 Then
        Delim = ","
    End If
    fT = dbRs.Fields.Count - 1
    For fC = 0 To fT
        fN = dbRs.Fields(fC).Name
        fL = StringsConcat(fL, fN, Delim)
    Next
    DaoFldNames = RemoveDelim(fL, Delim)
    Err.Clear
End Function
Public Sub LstViewSetMemory(lstReport As ListView, totRecords As Long)
    On Error Resume Next
    SendMessage lstReport.hWnd, LVM_SETITEMCOUNT, totRecords, 0&
    Err.Clear
End Sub
Function ConsolidateFiles(ByVal sDest As String, ByVal lstReport As ListView) As Boolean
    On Error Resume Next
    Dim bTemp() As Byte
    Dim nDestFile As Long
    Dim nSrcFile As Long
    Dim i As Long
    Dim N As Long
    Dim s As String
    N = lstReport.ListItems.Count
    ReDim sTemp(N)
    nDestFile = FreeFile
    Open sDest For Binary Access Write As nDestFile
        For i = 1 To N
            nSrcFile = FreeFile
            s = lstReport.ListItems(i).Text
            Open s For Binary Access Read As nSrcFile
                ReDim bTemp(LOF(nSrcFile) - 1)
                Get nSrcFile, , bTemp
                Put nDestFile, , bTemp
            Close nSrcFile
        Next
    Close nDestFile
    ConsolidateFiles = True
    Err.Clear
End Function
Function ConsolidateFilesCollection(ByVal sDest As String, ByVal lstCollection As Collection) As Boolean
    On Error Resume Next
    Dim bTemp() As Byte
    Dim nDestFile As Long
    Dim nSrcFile As Long
    Dim i As Long
    Dim N As Long
    Dim s As String
    N = lstCollection.Count
    ReDim sTemp(N)
    nDestFile = FreeFile
    Open sDest For Binary Access Write As nDestFile
        For i = 1 To N
            nSrcFile = FreeFile
            s = lstCollection(i)
            Open s For Binary Access Read As nSrcFile
                ReDim bTemp(LOF(nSrcFile) - 1)
                Get nSrcFile, , bTemp
                Put nDestFile, , bTemp
            Close nSrcFile
        Next
    Close nDestFile
    ConsolidateFilesCollection = True
    Err.Clear
End Function
Public Sub ImportSR0003(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim spLines() As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim xPart As New Collection
    Dim fPart As String
    Dim hrRs As New ADODB.Recordset
    Dim hrMonth As String
    Dim sDate As String
    Dim eDate As String
    hrMonth = FileToken(strFile, "fo")
    sDate = StartEndDate(hrMonth, "s")
    eDate = StartEndDate(hrMonth, "e")
    If IsDate(sDate) = False Or IsDate(eDate) = False Then
        Call MyPrompt("The file name should reflect the month of the rental in YYYYMM format.", "o", "e", "Import House Rent")
    Err.Clear
        Exit Sub
    End If
    Execute "delete from `house rent` where yyyymm = " & hrMonth
    Set hrRs = OpenRs("house rent")
    strLine = FileData(strFile)
    rsTot = StrParse(spLines, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing house rent..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(spLines(rsCnt))
        Select Case Left$(strLine, 1)
        Case "0"
        Case Else
            strLine = "0" & strLine
            fPart = StringPart(strLine, 1, " ")
            Select Case fPart
            Case "0-", "01", "0PPPPPPPPPPPP", "0PPP", "0MANAGER:", "0RECURRENT", "0-PREVIOUS", "0+", "0--"
            Case "0E-MAIL:", "0TEL:", "0FAX:", "0PRIVATE", "0MARSHALLTOWN", "0-NEW", "0-AMOUNT", "01**DJDE"
            Case "02107", "01CODE:", "0HOUSE", "01PERSAL", "0", "0SALARY", "0REPORT", "0RUN", "0PERSALNO", "0-----------"
            Case "0-----------------------------------------------------------------------------------------------------------------------------------"
            Case Else
                hrRs.AddNew
                hrRs.Fields("Persal") = Trim$(Mid$(strLine, 2, 8))
                hrRs.Fields("SurnameInitials") = Trim$(Mid$(strLine, 14, 27))
                hrRs.Fields("Paypoint") = Trim$(Mid$(strLine, 42, 8))
                hrRs.Fields("Reference") = Trim$(Mid$(strLine, 52, 17))
                hrRs.Fields("Reason") = Trim$(Mid$(strLine, 71, 6))
                hrRs.Fields("Notch") = Trim$(Mid$(strLine, 78, 10))
                hrRs.Fields("Amount") = Trim$(Mid$(strLine, 92, 11))
                hrRs.Fields("ArrearAmount") = Trim$(Mid$(strLine, 106, 13))
                hrRs.Fields("Balance") = Trim$(Mid$(strLine, 120))
                hrRs.Fields("RentalDate") = eDate
                hrRs.Fields("Yyyymm") = Val(hrMonth)
                UpdateRs hrRs
            End Select
        End Select
        DoEvents
    Next
    StatusMessage frmObj
    Err.Clear
End Sub
Public Sub TreeViewRemoveChildren(treeData As TreeView, ByVal ParentNodePos As Long)
    On Error Resume Next
    Dim nodeParent As Node
    Dim nodeChild As Node
    Dim nodeIdx As Long
    Dim nodeTot As Long
    Set nodeParent = treeData.Nodes(ParentNodePos)
    nodeTot = nodeParent.Children
    Do Until nodeTot = 0
        Set nodeChild = nodeParent.Child
        nodeIdx = nodeChild.Index
        If nodeIdx <> 0 Then
            treeData.Nodes.Remove nodeIdx
        End If
        nodeTot = treeData.Nodes(ParentNodePos).Children
        DoEvents
    Loop
    Set nodeParent = Nothing
    Set nodeChild = Nothing
    Err.Clear
End Sub
Function TreeViewAddPathWithKey(TreeV As TreeView, ByVal sPath As String, Optional ByVal Image As String = "", Optional ByVal SelectedImage As String = "", Optional ByVal Tag As String = "", Optional Delim As String = "\") As Long
    On Error GoTo ErrHandler
    Dim arrayPath() As String
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim sParent As String
    Dim nodeA As Node
    Dim pParent As String
    ' split the path to be subitems
    ArrayTot = StrParse(arrayPath, sPath, Delim)
    ' how many items do we have
    ArrayTot = UBound(arrayPath)
    For arrayCnt = 1 To ArrayTot
        ' get the current path
        sParent = MvFromMv(sPath, 1, arrayCnt, "\")
        Select Case arrayCnt
        Case 1
            ' we are adding the first node
            Set nodeA = TreeV.Nodes.Add(, , sParent, sParent, Image, SelectedImage)
        Case Else
            ' we are adding other nodes
            ' read the previous parent
            pParent = MvFromMv(sPath, 1, arrayCnt - 1, "\")
            Set nodeA = TreeV.Nodes.Add(pParent, tvwChild, sParent, arrayPath(arrayCnt), Image, SelectedImage)
        End Select
    Next
    TreeViewAddPathWithKey = nodeA.Index
    Err.Clear
    Exit Function
ErrHandler:
    If Err.Number = 35602 Then
        Set nodeA = TreeV.Nodes(sParent)
        Resume Next
    End If
    Err.Clear
End Function
Public Function FormatFileSize(sNumBytes As String) As String
    On Error Resume Next
    Dim NumBytes As Double
    Const SIZE_KB As Double = 1024
    Const SIZE_MB As Double = 1024 * SIZE_KB
    Const SIZE_GB As Double = 1024 * SIZE_MB
    Const SIZE_TB As Double = 1024 * SIZE_GB
    NumBytes = Val(sNumBytes)
    Select Case NumBytes
    Case Is < SIZE_KB
        FormatFileSize = Format$(NumBytes) & " bytes"
    Case Is < SIZE_MB
        FormatFileSize = Format$(NumBytes / SIZE_KB, "0.00") & " KB"
    Case Is < SIZE_GB
        FormatFileSize = Format$(NumBytes / SIZE_MB, "0.00") & " MB"
    Case Is < SIZE_TB
        FormatFileSize = Format$(NumBytes / SIZE_GB, "0.00") & " GB"
    Case Else
        FormatFileSize = Format$(NumBytes / SIZE_TB, "0.00") & " TB"
    End Select
    Err.Clear
End Function
Public Sub PutProgressBarInStatusBar(objStatusBar As Variant, objProgressBar As Variant, PnlNumber As Integer)
    On Error Resume Next
    Dim R As RECT
    SetParent objProgressBar.hWnd, objStatusBar.hWnd
    SendMessage objStatusBar.hWnd, SB_GETRECT, PnlNumber - 1, R
    MoveWindow objProgressBar.hWnd, R.Left + 1, R.Top + 1, R.Right - R.Left - 2, R.Bottom - R.Top - 2, True
    Err.Clear
End Sub
Public Sub ClosePeriod(frmObj As Form, ByVal sPeriod As String, Optional ByVal useItem As String, Optional DoReport As Boolean = False, Optional InitializePeriodItem As Boolean = False, Optional UseTb As Boolean = False, Optional InitializeFY As Boolean = False, Optional PromptUser As Boolean = True)
    On Error Resume Next
    frmObj.lstReport.Tag = "closeperiod"
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim tbS As New ADODB.Recordset
    Dim tbT As New ADODB.Recordset
    Dim sitem As String
    Dim sAmount As String
    Dim sKey As String
    Dim oAmount As String
    Dim spLine(1 To 4) As String
    Dim doCompile As Integer
    Dim lTot As Long
    Dim eDate As String
    Dim sDate As String
    Dim qry1 As String
    eDate = StartEndDate(sPeriod, "e")
    doCompile = vbYes
    If PromptUser = True Then
        doCompile = MyPrompt("Do you want to recompile the balances?", "yn", "q", "Confirm Recompile: " & YearMonthDesc(sPeriod))
    End If
    qrySql = "select Item,Amount from ledger where Yyyymm = '" & EscIn(sPeriod) & "'"
    If Len(useItem) > 0 Then
        qrySql = "select Item,Amount from ledger where Yyyymm = '" & EscIn(sPeriod) & "' and item = '" & EscIn(useItem) & "';"
    End If
    If InitializePeriodItem = True Then
        qrySql = "select Item,Amount from ledger where BasDate  <= '" & SwapDate(eDate) & "' and item = '" & EscIn(useItem) & "' and narration <> 'OPENING BALANCE';"
    End If
    If UseTb = True Then
        qrySql = "select Item,Amount from `trial balances` where period = " & sPeriod & ";"
    End If
    If InitializeFY = True Then
        qrySql = "select Item,Amount from ledger where BasDate  <= '" & SwapDate(eDate) & "' and narration <> 'OPENING BALANCE';"
    End If
    If doCompile = vbNo Then GoTo NextStep
    qry1 = "update `trial balances` set detailed = '0.00', balances = '0', OnTb = 'N' where period = '" & EscIn(sPeriod) & "';"
    If Len(useItem) > 0 Then
        qry1 = "update `trial balances` set detailed = '0.00', balances = '0', OnTb = 'N' where period = '" & EscIn(sPeriod) & "' and item = '" & EscIn(useItem) & "';"
    End If
    Execute qry1
    Set tbS = OpenRs(qrySql)
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Summing bas transactions..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sitem = MyRN(tbS.Fields("Item"))
        sAmount = ProperAmount(MyRN(tbS.Fields("Amount")))
        sKey = sPeriod & "*" & sitem
        Set tbT = SeekRs("ID", sKey, "Trial Balances")
        Select Case tbT.EOF
        Case True
            tbT.AddNew
            tbT.Fields("ID") = sKey
            tbT.Fields("Period") = sPeriod
            tbT.Fields("Item") = sitem
            tbT.Fields("Detailed") = sAmount
            tbT.Fields("OnTb") = "N"
            tbT.Fields("Balances") = "0"
            tbT.Fields("Date") = eDate
            UpdateRs tbT
        Case Else
            oAmount = ProperAmount(MyRN(tbT.Fields("Detailed")))
            oAmount = Val(oAmount) + Val(ProperAmount(sAmount))
            oAmount = ProperAmount(oAmount)
            tbT.Fields("Detailed") = oAmount
            UpdateRs tbT
        End Select
        DoEvents
        tbS.MoveNext
    Next
    ProgBarClose frmObj.progBar
NextStep:
    Execute "update `trial balances` set balances = '1', ontb = 'Y' where period = " & EscIn(sPeriod) & " and amount = detailed;"
    Set tbS = OpenRs("select * from `Trial Balances` where period = " & sPeriod)
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Updating the opening balances..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sitem = MyRN(tbS.Fields("Item"))
        sAmount = MyRN(tbS.Fields("Amount"))
        sKey = MyRN(tbS.Fields("balances"))
        sDate = MyRN(tbS.Fields("Date"))
        sDate = DateAdd("d", 1, sDate)
        If sAmount = "0.00" Or sKey = "0" Then GoTo NextRecord
        qrySql = "select * from ledger where item = '" & EscIn(sitem) & "' and BasDate = '" & SwapDate(sDate) & "' and narration = 'OPENING BALANCE'"
        Set tbT = OpenRs(qrySql)
        If tbT.EOF = True Then
            tbT.AddNew
            tbT.Fields("Item") = sitem
            tbT.Fields("entrytype") = "OB"
            tbT.Fields("funcArea") = ""
            tbT.Fields("narration") = "OPENING BALANCE"
            tbT.Fields("Responsibility") = ""
            tbT.Fields("Objective") = ""
            tbT.Fields("Project") = ""
            tbT.Fields("Reference") = ""
            tbT.Fields("BasAudit") = ""
            tbT.Fields("User") = ""
            tbT.Fields("basdate") = sDate
            tbT.Fields("Amount") = sAmount
            tbT.Fields("Yyyymm") = Format$(sDate, "YYYYMM")
            tbT.Fields("district") = "HO"
            tbT.Fields("Beneficiary") = ""
            tbT.Fields("persal") = ""
            UpdateRs tbT
        Else
            tbT.Fields("entrytype") = "OB"
            tbT.Fields("Amount") = sAmount
            UpdateRs tbT
        End If
        DoEvents
NextRecord:
        tbS.MoveNext
    Next
    ProgBarClose frmObj.progBar
    Set tbS = Nothing
    Set tbT = Nothing
    If DoReport = True Then
        frmObj.Caption = "eFas - Trial Balance Check Report For " & sPeriod
        ViewSQLNew "select parent,item,amount,detailed from `trial balances` where period = " & sPeriod & " and balances = '0' and amount <> '0.00' order by parent,item", frmObj.lstReport, "Economic Classification,Account,Trial Balance,Ledger", , , , , , True, , , "amount,detailed"
        LstViewSumColumns frmObj.lstReport, True, "Trial Balance", "Ledger"
        frmObj.lstReport.Tag = "y"
        ResetFilter frmObj.lstReport, frmObj.lstValue, frmObj.cboField, frmObj.chkRemove
        ZeroAccounts frmObj, sPeriod
    End If
    StatusMessage frmObj
    Beep
    Err.Clear
End Sub
Public Sub ZeroAccounts(frmObj As Form, ByVal strMonth As String)
    On Error Resume Next
    Dim qrySql As String
    qrySql = "select Item,Amount from `Trial Balances` where period = '" & EscIn(strMonth) & "' and OnTb = 'N' order by Item;"
    ViewSQLNew qrySql, frmObj.lstAllocations1, "Account,Amount", , , , , False, True, , , "Amount"
    LstViewSumColumns frmObj.lstAllocations1, True, "Amount"
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, "Accounts Not In Trial Balance - " & strMonth, frmObj.lstAllocations1, , False, True, False
    Err.Clear
End Sub
Sub ShowTb(frmObj As Form, qrySql As String)
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine(1 To 4) As String
    Dim sAmount As String
    Set tb = OpenRs(qrySql)
    rsTot = AffectedRecords
    LstViewSetMemory frmObj.lstReport, rsTot
    For rsCnt = 1 To rsTot
        spLine(1) = MyRN(tb.Fields("parent"))
        spLine(2) = MyRN(tb.Fields("item"))
        sAmount = ProperAmount(MyRN(tb.Fields("amount")))
        If Val(sAmount) > 0 Then
            spLine(3) = MakeMoney(sAmount)
            spLine(4) = "0.00"
        Else
            spLine(4) = MakeMoney(Replace$(sAmount, "-", ""))
            spLine(3) = "0.00"
        End If
        Call LstViewUpdate(spLine, frmObj.lstReport, "")
        tb.MoveNext
    Next
    tb.Close
    Set tb = Nothing
    Err.Clear
End Sub
Sub ImportFunds(frmObj As Form)
    On Error Resume Next
    Dim cTotal As Long
    Dim cCount As Long
    Dim strFile As String
    resp = MyPrompt("Before you can import the code structure..." & vbCr & "please ensure that the file names excluding the extensions are the follow:" & vbCr & "funds or items or objectives or responsibilities or projects." & vbCr & vbCr & "Continue?", "yn", "q", "Import Bas Structure")
    If resp = vbNo Then Exit Sub
    strFile = BrowseForFolder(frmObj.hWnd, "Import Bas Code Structure")
    If Len(strFile) = 0 Then Exit Sub
    LstViewFromComputerFolder frmObj.lstReport, strFile, "Bas Code Structure Files", "*.txt"
    StatusMessage frmObj, frmObj.lstReport.ListItems.Count & " file(s) selected."
    frmObj.lstReport.Refresh
    resp = MyPrompt("You are about to import " & frmObj.lstReport.ListItems.Count & " code structure files. Are you sure?", "yn", "q", "Import Code Structure")
    If resp = vbNo Then Exit Sub
    cTotal = frmObj.lstReport.ListItems.Count
    For cCount = 1 To cTotal
        strFile = frmObj.lstReport.ListItems(cCount).Text
        Select Case LCase$(FileToken(strFile, "fo"))
        Case "funds", "items", "objectives", "responsibilities", "projects"
            UploadCodeStructure frmObj, strFile, True
        End Select
    Next
    Err.Clear
End Sub
Public Sub CodeStructureSearch(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim askCode As String
    askCode = InputBox("Please enter the search string to search for:", "Search Code Structure")
    If Len(askCode) = 0 Then Exit Sub
    Select Case LCase$(strFile)
    Case "items", "objectives", "responsibilities"
    Case Else
    Err.Clear
        Exit Sub
    End Select
    frmObj.Caption = "eFas - " & ProperCase(strFile & " Search")
    qrySql = "select Code,Description,PostingLevel,Path from " & strFile & " where " & BuildSQL("ado", "Path", askCode) & "order by Code;"
    ViewSQLNew qrySql, frmObj.lstReport, "ID,Description,Posting Level,Path"
    StatusMessage frmObj, frmObj.lstReport.ListItems.Count & " items found."
    ResetFilter frmObj.lstReport, frmObj.lstValue, frmObj.cboField, frmObj.chkRemove
    Err.Clear
End Sub
Public Function BasPersalItems(ByVal sDate As String, ByVal eDate As String) As String
    On Error Resume Next
    Dim persalItems As String
    Dim basItems As String
    Dim sql As String
    ' read the persal items available
    sql = "select distinct item from persal where paydate >= '" & SwapDate(sDate) & "' and paydate <= '" & SwapDate(eDate) & "' order by item;"
    persalItems = DistinctColumnString(sql, "item", VM)
    ' read the bas items available based on interface transactions
    sql = "select distinct item from ledger where BasDate >= '" & SwapDate(sDate) & "' and basdate <= '" & SwapDate(eDate) & "' and user = 'IFBS11BS' order by item;"
    basItems = DistinctColumnString(sql, "item", VM)
    ' consolidate the items
    basItems = basItems & VM & persalItems
    basItems = MvRemoveDuplicates(basItems, VM)
    BasPersalItems = MvRemoveBlanks(basItems, VM)
    Err.Clear
End Function
Public Sub ProgBarInit(progBar As ProgressBar, totItems As Long)
    On Error Resume Next
    frmLogin.CheckUpdate.Enabled = False
    frmLogin.IdleMonitor.Enable = False
    ProgBarClose progBar
    progBar.Max = totItems
    progBar.Min = 0
    Err.Clear
End Sub
Public Sub ProgBarClose(progBar As ProgressBar)
    On Error Resume Next
    progBar.Value = 0
    frmLogin.CheckUpdate.Enabled = True
    frmLogin.IdleMonitor.Enable = True
    Err.Clear
End Sub
Public Sub TreeViewNodeOperation(treeData As TreeView, ByVal sOperation As String)
    On Error Resume Next
    Dim nodeCnt As Long
    Dim nodeTot As Long
    Dim sNode As Variant
    nodeTot = treeData.Nodes.Count
    If nodeTot = 0 Then Exit Sub
    Select Case LCase$(Left$(sOperation, 1))
    Case "e"
        For nodeCnt = nodeTot To 1 Step -1
            If treeData.Nodes(nodeCnt).Expanded = False Then
                treeData.Nodes(nodeCnt).Expanded = True
            End If
        Next
    Case "c"
        For nodeCnt = 1 To nodeTot
            If treeData.Nodes(nodeCnt).Expanded = True Then
                treeData.Nodes(nodeCnt).Expanded = False
            End If
        Next
    End Select
    If treeData.Nodes.Count >= 1 Then
        treeData.Nodes(1).EnsureVisible
    End If
    Set sNode = Nothing
    Err.Clear
End Sub
Public Sub DeBold(objLabel As Label, Optional MakeBold As Boolean = False)
    On Error Resume Next
    ' change font to regular
    objLabel.Font.Bold = MakeBold
    objLabel.ForeColor = &HFF0000
    objLabel.Refresh
    Err.Clear
End Sub
Public Sub Iconize(objForm As Form, objControl As Label, Optional colForeColor As VBRUN.ColorConstants = vbBlue)
    On Error Resume Next
    ' change icon of label on mouse move
    With objControl
        .Font.Bold = True
        .ForeColor = colForeColor
        SetHandCur objForm, True
        .MousePointer = 99
    End With
    Err.Clear
End Sub
' Desc: Get the Hand Cursor
Public Sub SetHandCur(objForm As Form, Hand As Boolean)
    On Error Resume Next
    If Hand = True Then
        objForm.MousePointer = 99
    Else
        objForm.MousePointer = 0
    End If
    Err.Clear
End Sub
Public Sub TreeView_ChildrenToCollection(treeData As Variant, ByVal Treenodepos As Long, newCol As Collection)
    On Error Resume Next
    Dim nodeParent As Variant
    Dim nodeChild As Variant
    Dim nodeIdx As Long
    Dim nodeTot As Long
    Dim nodeCnt As Long
    nodeCnt = 0
    Set newCol = New Collection
    Set nodeParent = treeData.Nodes(Treenodepos)
    nodeTot = nodeParent.Children
    If nodeTot = 0 Then Exit Sub
    Set nodeChild = nodeParent.Child
    Do Until nodeChild Is Nothing
        newCol.Add nodeChild.FullPath
        Set nodeChild = nodeChild.Next
    Loop
    Set nodeParent = Nothing
    Set nodeChild = Nothing
    Err.Clear
End Sub
Public Sub AmountsBasedOnFunctions(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = 1 To rsTot
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If LCase$(spLines(5)) = "no" Then spLines(9) = "0.00"
        If LCase$(spLines(6)) = "no" Then spLines(10) = "0.00"
        If LCase$(spLines(7)) = "no" Then spLines(11) = "0.00"
        ' recalculate total budget to be transferred
        spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
        spLines(12) = MakeMoney(spLines(12))
        If UBound(spLines) >= 19 Then
            ' recalculate actual
            spLines(18) = Val(ProperAmount(spLines(13))) + Val(ProperAmount(spLines(14))) + Val(ProperAmount(spLines(15))) + Val(ProperAmount(spLines(16))) + Val(ProperAmount(spLines(17)))
            spLines(18) = MakeMoney(spLines(18))
            ' calculate the variance between budgeted transfers and actual transfers
            spLines(19) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(18)))
            spLines(19) = MakeMoney(spLines(19))
        End If
        'update report
        Call LstViewUpdate(spLines, lstReport, CStr(rsCnt))
        lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
    Next
    Err.Clear
End Sub
Public Sub AmountsBasedOnFunctionsIDS(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = 1 To rsTot
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If LCase$(spLines(4)) = "no" Then spLines(8) = "0.00"
        If LCase$(spLines(5)) = "no" Then spLines(9) = "0.00"
        If LCase$(spLines(6)) = "no" Then spLines(10) = "0.00"
        ' recalculate total budget to be transferred
        spLines(12) = Val(ProperAmount(spLines(8))) + Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
        spLines(12) = MakeMoney(spLines(12))
        If UBound(spLines) >= 20 Then
            ' recalculate actual
            spLines(19) = Val(ProperAmount(spLines(13))) + Val(ProperAmount(spLines(14))) + Val(ProperAmount(spLines(15))) + Val(ProperAmount(spLines(16))) + Val(ProperAmount(spLines(17))) + Val(ProperAmount(spLines(18)))
            spLines(19) = MakeMoney(spLines(19))
            ' calculate the variance between budgeted transfers and actual transfers
            spLines(20) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(19)))
            spLines(20) = MakeMoney(spLines(20))
        End If
        'update report
        Call LstViewUpdate(spLines, lstReport, CStr(rsCnt))
        lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
    Next
    Err.Clear
End Sub
Public Sub RttAndLsm(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If ProperAmount(spLines(14)) <> "0.00" And ProperAmount(spLines(18)) <> "0.00" Then
        Else
            lstReport.ListItems.Remove rsCnt
        End If
NextLine:
    Next
    Err.Clear
End Sub
Public Sub Rtt(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If ProperAmount(spLines(18)) <> "0.00" Then
        Else
            lstReport.ListItems.Remove rsCnt
        End If
NextLine:
    Next
    Err.Clear
End Sub
Sub ForwardPayments(frmObj As Form, ByVal sPayments As String, ByVal sFrom As String, ByVal sTo As String, ByVal sDate As String)
    On Error Resume Next
    If Len(sPayments) = 0 Or Len(sFrom) = 0 Or Len(sTo) = 0 Or Len(sDate) = 0 Then
        Call MyPrompt("The process to forward the payment(s) could not be started, there is an error reading the payments, the sender, the recipient or the date, please retry.", "o", "e", "Forward Error")
    Err.Clear
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim xPay As String
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sPay As String
    Dim issError As New Collection
    Dim tb As New ADODB.Recordset
    Dim sData As String
    xPay = Replace$(sPayments, vbNewLine, ",")
    xPay = Replace$(sPayments, "'", "")
    xPay = Replace$(xPay, ".", "")
    xPay = MvRemoveBlanks(xPay, ",")
    xPay = MvRemoveDuplicates(xPay, ",")
    spTot = StrParse(spLine, xPay, ",")
    ProgBarInit frmObj.progBar, spTot
    StatusMessage frmObj, "Fowarding payments, please be patient..."
    For spCnt = 1 To spTot
        frmObj.progBar.Value = spCnt
        sPay = Trim$(ExtractNumbers(spLine(spCnt)))
        If Len(sPay) > 0 Then
            If RecordExists("Issued Payment Advices", "Sequence", sPay) = True Then
                ' update the payment advice itself
                sData = sTo & VM & sDate & VM & usrFullName & VM & Format$(Now, "dd/mm/yyyy hh:mm:ss ampm")
                WriteRecordMv "Payment advice", "serial", sPay, "Currentlocation,DateSent,EditedByUser,DateEdited", sData, VM
                'RecycleRecord "Payment Advice", "Serial", sPay, "Forward Payment Advice " & sPay & " From " & usrFullName & " To " & sTo, usrFullName, "Audit Trail"
                ' update the history file
                Set tb = OpenRs("history", , 1)
                tb.AddNew
                tb.Fields("Payment") = Val(sPay)
                tb.Fields("Sender") = sFrom
                tb.Fields("DateRecorded") = Now()
                tb.Fields("Recipient") = sTo
                tb.Fields("dateSent") = sDate
                UpdateRs tb
            Else
                issError.Add sPay
            End If
        End If
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    CurrentLocation
    StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " items listed."
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Public Sub Pass_WordCreateMemo(sOffice As String, sPayment As String, sDescription As String, sSupplier As String, sAmount As String, sUser As String, sUserTelephone As String, sUserEmail As String, sAllocations As String, Optional ByVal sHeader As String = "Passed Q.A.")
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim wordDoc As New Word.Document
    Dim wordFile As String
    Dim wordNew As String
    'bFile = ""
    wordFile = App.Path & "\Templates\pass.doc"
    If FileExists(wordFile) = False Then
        Call MyPrompt("The pass template file does not exist, please re-install the application.", "o", "e", "Pass Template File")
    Err.Clear
        Exit Sub
    End If
    wordNew = App.Path & "\Documents\Payment Advice " & sPayment & ".doc"
    If FileExists(wordNew) = True Then
        Kill wordNew
    End If
    Do Until FileExists(wordNew) = True
        DoEvents
        FileCopy wordFile, wordNew
    Loop
    Set wordDoc = wordDoc.Application.Documents.Open(wordNew)
    wordDoc.ActiveWindow.WindowState = wdWindowStateMinimize
    wordDoc.Activate
    If wordDoc.Application.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        wordDoc.Application.ActiveWindow.Panes(2).Close
    End If
    If wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdNormalView Or wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    wordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    wordDoc.Application.Selection.HeaderFooter.Shapes("WordArt 2").Select
    wordDoc.Application.Selection.ShapeRange.TextEffect.Text = sHeader
    wordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    WordReplace wordDoc, "<userfullname>", sUser
    WordReplace wordDoc, "<usertelephone>", sUserTelephone
    WordReplace wordDoc, "<useremail>", sUserEmail
    WordReplace wordDoc, "<payment>", sPayment
    WordReplace wordDoc, "<description>", sDescription
    WordReplace wordDoc, "<office>", sOffice
    WordReplace wordDoc, "<supplier>", sSupplier
    WordReplace wordDoc, "<bank>", frmQA.txtBankAccNo.Text
    WordReplace wordDoc, "<branch>", frmQA.txtBranchCode.Text
    WordReplace wordDoc, "<type>", frmQA.cboAccountType.Text
    WordReplace wordDoc, "<amount>", sAmount
    WordReplace wordDoc, "<date>", Format$(Now(), "dd/mm/yyyy")
    sAllocations = FormatTextForWord(sAllocations, 2)
    WordFindInsert wordDoc, "<allocations>", sAllocations
    WordReplace wordDoc, "<allocations>", ""
    Screen.MousePointer = vbDefault
    wordDoc.Save
    wordDoc.Close
    wordDoc.Application.Quit
    If FileExists(wordNew) = False Then
        resp = MyPrompt("File Name:" & vbCr & wordNew & vbCr & vbCr & "Operation - Failure.", "o", "c", "File Error")
    Err.Clear
        Exit Sub
    End If
    ViewFile wordNew, "print"
    Err.Clear
End Sub
Public Sub Fail_WordCreateMemo(sOffice As String, sPayment As String, sDescription As String, sSupplier As String, sAmount As String, sUser As String, sUserTelephone As String, sUserEmail As String, sExceptions As String, sOtherExceptions As String, Optional ByVal sHeader As String = "Failed Q.A.")
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim wordDoc As New Word.Document
    Dim wordFile As String
    Dim wordNew As String
    'bFile = ""
    wordFile = App.Path & "\Templates\fail.doc"
    If FileExists(wordFile) = False Then
        Call MyPrompt("The fail template file does not exist, please re-install the application.", "o", "e", "Fail Template File")
    Err.Clear
        Exit Sub
    End If
    wordNew = App.Path & "\Documents\Payment Advice " & sPayment & ".doc"
    If FileExists(wordNew) = True Then
        Kill wordNew
    End If
    Do Until FileExists(wordNew) = True
        DoEvents
        FileCopy wordFile, wordNew
    Loop
    Set wordDoc = wordDoc.Application.Documents.Open(wordNew)
    wordDoc.ActiveWindow.WindowState = wdWindowStateMinimize
    wordDoc.Activate
    If wordDoc.Application.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        wordDoc.Application.ActiveWindow.Panes(2).Close
    End If
    If wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdNormalView Or wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        wordDoc.Application.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    wordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    wordDoc.Application.Selection.HeaderFooter.Shapes("WordArt 2").Select
    wordDoc.Application.Selection.ShapeRange.TextEffect.Text = sHeader
    wordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    WordReplace wordDoc, "<userfullname>", sUser
    WordReplace wordDoc, "<usertelephone>", sUserTelephone
    WordReplace wordDoc, "<useremail>", sUserEmail
    WordReplace wordDoc, "<payment>", sPayment
    WordReplace wordDoc, "<office>", sOffice
    WordReplace wordDoc, "<description>", sDescription
    WordReplace wordDoc, "<supplier>", sSupplier
    WordReplace wordDoc, "<amount>", sAmount
    WordReplace wordDoc, "<date>", Format$(Now(), "dd/mm/yyyy")
    sExceptions = FormatTextForWord(sExceptions, 0)
    WordFindInsert wordDoc, "<exceptions>", Replace$(sExceptions, vbNewLine, vbCr)
    WordReplace wordDoc, "<exceptions>", ""
    sOtherExceptions = FormatTextForWord(sOtherExceptions, 0)
    WordFindInsert wordDoc, "<other>", Replace$(sOtherExceptions, vbNewLine, vbCr)
    WordReplace wordDoc, "<other>", ""
    Screen.MousePointer = vbDefault
    wordDoc.Save
    wordDoc.Close
    wordDoc.Application.Quit
    If FileExists(wordNew) = False Then
        resp = MyPrompt("File Name:" & vbCr & wordNew & vbCr & vbCr & "Operation - Failure.", "o", "c", "File Error")
    Err.Clear
        Exit Sub
    End If
    ViewFile wordNew, "print"
    Err.Clear
End Sub
Public Function ImportSchoolBudget(frmObj As Form, StrSource As String) As Boolean
    On Error Resume Next
    Dim dbs As DAO.Database
    Dim tbS As DAO.Recordset
    Dim tbT As New ADODB.Recordset
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim sSection21 As String
    Dim sFuncA As String
    Dim sFuncC As String
    Dim sFuncD As String
    Dim sYear As String
    Dim sId As String
    Dim sTotal As String
    If DaoTableExists(StrSource, "schoolbudget") = False Then
        ImportSchoolBudget = False
    Err.Clear
        Exit Function
    End If
    Set dbs = DAO.OpenDatabase(StrSource)
    Set tbS = dbs.OpenRecordset("schoolbudget")
    sYear = FileToken(StrSource, "fo")
    rsTot = tbS.RecordCount
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing the school budget, please be patient..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sId = RN(tbS.Fields("ref")) & "-" & sYear
        Set tbT = SeekRs("ID", sId, "Schools Budget")
        If tbT.EOF = True Then tbT.AddNew
        sSection21 = RN(tbS.Fields("section21"))
        sFuncA = Trim$(RN(tbS.Fields("FuncA")))
        sFuncC = Trim$(RN(tbS.Fields("FuncC")))
        sFuncD = Trim$(RN(tbS.Fields("FuncD")))
        If sSection21 = "0" Then sSection21 = "No"
        If sSection21 = "-1" Then sSection21 = "Yes"
        If sSection21 = "False" Then sSection21 = "No"
        If sSection21 = "True" Then sSection21 = "Yes"
        If sFuncA = "0" Then sFuncA = "No"
        If sFuncA = "-1" Then sFuncA = "Yes"
        If sFuncA = "False" Then sFuncA = "No"
        If sFuncA = "True" Then sFuncA = "Yes"
        If sFuncC = "0" Then sFuncC = "No"
        If sFuncC = "-1" Then sFuncC = "Yes"
        If sFuncC = "False" Then sFuncC = "No"
        If sFuncC = "True" Then sFuncC = "Yes"
        If sFuncD = "0" Then sFuncD = "No"
        If sFuncD = "-1" Then sFuncD = "Yes"
        If sFuncD = "False" Then sFuncD = "No"
        If sFuncD = "True" Then sFuncD = "Yes"
        tbT.Fields("ID") = sId
        tbT.Fields("Year") = sYear
        tbT.Fields("EMIS") = RN(tbS.Fields("ref"))
        tbT.Fields("district") = RN(tbS.Fields("DIS"))
        tbT.Fields("schoolname") = ProperCase(RN(tbS.Fields("School Name")))
        tbT.Fields("section21") = sSection21
        tbT.Fields("lsm") = ProperAmount(Round(RN(tbS.Fields("lsm")), 2))
        tbT.Fields("SERVICES") = ProperAmount(Round(RN(tbS.Fields("SERVICES")), 2))
        tbT.Fields("Maintenance") = ProperAmount(Round(RN(tbS.Fields("Maintenance")), 2))
        tbT.Fields("FuncA") = sFuncA
        tbT.Fields("FuncC") = sFuncC
        tbT.Fields("FuncD") = sFuncD
        sTotal = Val(ProperAmount(MyRN(tbT.Fields("lsm")))) + Val(ProperAmount(MyRN(tbT.Fields("SERVICES")))) + Val(ProperAmount(MyRN(tbT.Fields("Maintenance"))))
        sTotal = ProperAmount(sTotal)
        tbT.Fields("Total") = sTotal
        UpdateRs tbT
        DoEvents
        tbS.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tbS.Close
    dbs.Close
    ImportSchoolBudget = True
    Set tbT = Nothing
    Set tbS = Nothing
    Set dbs = Nothing
    Err.Clear
End Function
Public Sub ImportEbtPayments(frmObj As Form, strFile As String, FileType As Ebt_Type)
    On Error Resume Next
    Dim tmpFile As String
    Dim intFile As Long
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim DisbursementDate As String
    Dim DisbursementNumber As String
    Dim Beneficiary As String
    Dim BankName As String
    Dim BranchName As String
    Dim AccountNumber As String
    Dim AccountType As String
    Dim AuthorizedBy As String
    Dim Payment1 As String
    Dim Payment2 As String
    Dim Status As String
    Dim SerialNumber As String
    Dim Amount As String
    Dim Account As String
    Dim tb As New ADODB.Recordset
    fName = FileToken(strFile, "fo")
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    Select Case FileType
    Case BasEbt
        Execute "delete from `ebt payments` where Yyyymm = " & fName & " and EbtType = 'Bas';"
    Case SapEbt
        Execute "delete from `ebt payments` where Yyyymm = " & fName & " and EbtType = 'Sap';"
    End Select
    Set tb = OpenRs("Ebt Payments", , 1)
    StatusMessage frmObj, "Importing ebt for " & fName
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        fPart = MvField(Trim$(strLine), 1, " ")
        Select Case fPart
        Case "CURRENT", "SAVINGS", "TRANSMISSION", "BOND", "SUBSCRIPTION"
            '  the contents of this line referes to the previous details, thus update those
            Account = Trim$(Mid$(fData(rsCnt), 22, 32))
            AccountType = StringPart(Account, 1, ":")
            AccountNumber = StringPart(Account, 2, ":")
            BranchName = Trim$(Mid$(fData(rsCnt), 55, 31))
            Payment2 = Trim$(Mid$(fData(rsCnt), 98, 7))
            SerialNumber = Trim$(Mid$(fData(rsCnt), 106, 10))
            AccountType = Trim$(Replace$(AccountType, "ACCOUNT", ""))
            BranchName = Trim$(Replace$(BranchName, "*", ""))
            tb.AddNew
            tb.Fields("DisbursementDate") = DisbursementDate
            tb.Fields("DisbursementNo") = Val(DisbursementNumber)
            tb.Fields("Beneficiary") = Beneficiary
            tb.Fields("BankName") = BankName
            tb.Fields("BranchName") = BranchName
            tb.Fields("AccountNo") = AccountNumber
            tb.Fields("AccountType") = AccountType
            tb.Fields("AuthorizedBy") = AuthorizedBy
            tb.Fields("PaymentNo") = Val(Payment2)
            tb.Fields("paymenttype") = Payment1
            tb.Fields("Status") = Status
            tb.Fields("Amount") = Amount
            tb.Fields("Yyyymm") = Format$(DisbursementDate, "yyyymm")
            Select Case FileType
            Case BasEbt
                tb.Fields("ebttype") = "Bas"
            Case PerEbt
                tb.Fields("ebttype") = "Per"
            Case SapEbt
                tb.Fields("ebttype") = "Sap"
            End Select
            UpdateRs tb
        Case Else
            If IsDate(fPart) = True Then
                DisbursementDate = Mid$(strLine, 1, 12)
                DisbursementNumber = Trim$(Mid$(strLine, 14, 7))
                Beneficiary = Trim$(Mid$(strLine, 22, 32))
                BankName = Trim$(Mid$(strLine, 55, 31))
                AuthorizedBy = Trim$(Mid$(strLine, 87, 10))
                Payment1 = Trim$(Mid$(strLine, 98, 7))
                Status = Trim$(Mid$(strLine, 106, 10))
                Amount = ProperAmount(Trim$(Mid$(strLine, 117)))
            End If
        End Select
NextLine:
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tb.Close
    Err.Clear
End Sub
Public Function PreprocessEbtPayments(frmObj As Form, strFile As String) As String
    On Error Resume Next
    Dim tmpFile As String
    Dim intFile As Long
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tFile As String
    Dim fFree As Long
    Dim oFree As Long
    Dim strOr As String
    fName = FileToken(strFile, "fo")
    tFile = App.Path & "\" & fName & ".txt"
    If FileExists(tFile) = True Then Kill tFile
    rsTot = FileLen(strFile)
    StatusMessage frmObj, "Pre-processing ebt payments file..."
    fFree = FreeFile
    Open tFile For Output Access Write As #fFree
        oFree = FreeFile
        Open strFile For Input Access Read As #oFree
            Do Until EOF(oFree)
                Line Input #oFree, strLine
                strOr = strLine
                strLine = Trim$(strLine)
                If Len(strLine) = 0 Then GoTo NextLine
                fPart = UCase$(MvField(Trim$(strLine), 1, " "))
                Select Case fPart
                Case "CURRENT", "SAVINGS", "TRANSMISSION", "BOND", "SUBSCRIPTION"
                    Print #fFree, strOr
                Case Else
                    If IsDate(fPart) = True Then
                        Print #fFree, strOr
                    End If
                End Select
NextLine:
            Loop
        Close #fFree
    Close #oFree
    PreprocessEbtPayments = tFile
    StatusMessage frmObj
    Err.Clear
End Function
Public Sub ImportPersalEbtPayments(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim DisbursementDate As String
    Dim DisbursementNumber As String
    Dim Payment As String
    Dim Status As String
    Dim SerialNumber As String
    Dim Amount As String
    Dim tb As New ADODB.Recordset
    fName = FileToken(strFile, "fo")
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    Execute "delete from `ebt payments` where Yyyymm = " & fName & " and EbtType = 'persal';"
    Set tb = OpenRs("Ebt Payments")
    StatusMessage frmObj, "Importing persal ebt for " & fName
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        fPart = StringPart(strLine, 1, " ")
        Select Case IsDate(fPart)
        Case True
            DisbursementDate = Trim$(Mid$(strLine, 1, 11))
            DisbursementNumber = Trim$(Mid$(strLine, 18, 9))
            Payment = Trim$(Mid$(strLine, 92, 8))
            Status = Trim$(Mid$(strLine, 103, 6))
            SerialNumber = Trim$(Mid$(strLine, 72, 11))
            Amount = Trim$(Mid$(strLine, 111))
            Amount = ProperAmount(Amount)
            tb.AddNew
            tb.Fields("DisbursementDate") = DisbursementDate
            tb.Fields("DisbursementNo") = Val(DisbursementNumber)
            tb.Fields("AuthorizedBy") = "APPER"
            tb.Fields("PaymentNo") = Val(Payment)
            tb.Fields("Status") = Status
            tb.Fields("Amount") = Amount
            tb.Fields("Yyyymm") = Format$(DisbursementDate, "yyyymm")
            tb.Fields("paymenttype") = "AP"
            tb.Fields("ebttype") = "Persal"
            UpdateRs tb
        End Select
NextLine:
    Next
    tb.Close
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Err.Clear
End Sub
Sub CurrentLocation()
    On Error Resume Next
    Dim spSup() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    frmQA.xpQA.ClearGroup "paymentmovement"
    frmQA.xpQA.DisableUpdates
    qrySql = "select distinct CurrentLocation from `Payment Advice` order by CurrentLocation;"
    spSup = DistinctColumnArray(qrySql, "CurrentLocation", True)
    spTot = UBound(spSup)
    For spCnt = 1 To spTot
        spStr = spSup(spCnt)
        frmQA.xpQA.AddItem "paymentmovement", "movement," & spStr, spStr, 1
    Next
    frmQA.xpQA.DisableUpdates False
    Err.Clear
End Sub
Public Sub ImportDRFile(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim DisbursementDate As String
    Dim DisbursementNumber As String
    Dim Beneficiary As String
    Dim Status As String
    Dim Amount As String
    Dim tb As New ADODB.Recordset
    Dim myRecs As String
    Dim recsTot As Long
    Dim recsCnt As Long
    Dim SerialNumber As String
    Dim sPayMethod As String
    fName = FileToken(strFile, "fo")
    StatusMessage frmObj, "Reading file contents " & fName
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    Execute "delete from `Disbursement Register` where yyyymm = " & fName
    Set tb = OpenRs("Disbursement Register")
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing file " & fName
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        fPart = Left$(strLine, 1)
        Select Case fPart
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            Select Case Mid$(strLine, 2, 1)
            Case "."
            Case Else
                DisbursementDate = Trim$(Mid$(strLine, 49, 10))
                DisbursementNumber = Val(Trim$(Mid$(strLine, 1, 10)))
                Status = Trim$(Mid$(strLine, 109))
                Amount = Trim$(Mid$(strLine, 84, 22))
                Amount = ProperAmount(Amount)
                Beneficiary = Trim$(Mid$(strLine, 14, 32))
                SerialNumber = Trim$(Mid$(strLine, 71, 10))
                sPayMethod = Trim$(Mid$(strLine, 62, 6))
                If IsDate(DisbursementDate) = False Then GoTo NextLine
                tb.AddNew
                tb.Fields("DisbursementNo") = Val(DisbursementNumber)
                tb.Fields("PayeeName") = Beneficiary
                tb.Fields("PaymentDate") = DisbursementDate
                tb.Fields("PaymentMethod") = sPayMethod
                tb.Fields("MicrNo") = Val(SerialNumber)
                tb.Fields("Amount") = Amount
                tb.Fields("Status") = Status
                tb.Fields("Yyyymm") = Val(Format$(DisbursementDate, "yyyymm"))
                UpdateRs tb
            End Select
        End Select
NextLine:
        DoEvents
    Next
    tb.Close
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Err.Clear
End Sub
Public Sub ImportRegisterOfDeposits(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim sDepositNo As String
    Dim sDepositStatus As String
    Dim sDepositDate As String
    Dim sYyyymm As String
    Dim sUserId As String
    Dim sBatchNo As String
    Dim sBatchStatus As String
    Dim sJournalAmount As String
    Dim sAmountDeposited As String
    Dim sTransNo As String
    fName = FileToken(strFile, "fo")
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing register of deposits..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        fPart = StringPart(strLine, 1, " ")
        Select Case fPart
        Case "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.", "11."
        Case Else
            If IsNumeric(fPart) = True Then
                If InStr(1, fPart, ",") = 0 Then
                    sDepositNo = Val(Trim$(Mid$(strLine, 1, 11)))
                    sDepositStatus = Trim$(Mid$(strLine, 14, 7))
                    sDepositDate = Trim$(Mid$(strLine, 23, 10))
                    sYyyymm = Format$(sDepositDate, "yyyymm")
                    sUserId = Trim$(Mid$(strLine, 46, 12))
                    sBatchNo = Val(Trim$(Mid$(strLine, 60, 10)))
                    sBatchStatus = Trim$(Mid$(strLine, 72, 6))
                    sJournalAmount = ProperAmount(Trim$(Mid$(strLine, 80, 22)))
                    sAmountDeposited = ProperAmount(Trim$(Mid$(strLine, 104)))
                    sTransNo = Val(Trim$(Mid$(strLine, 35, 9)))
                    If IsDate(sDepositDate) = True Then
                        Set tb = SeekRs("DepositNo", sDepositNo, "Register Of Deposits")
                        If tb.EOF = True Then tb.AddNew
                        tb.Fields("DepositNo") = Val(sDepositNo)
                        tb.Fields("DepositStatus") = sDepositStatus
                        tb.Fields("DepositDate") = sDepositDate
                        tb.Fields("Yyyymm") = Val(sYyyymm)
                        tb.Fields("UserId") = sUserId
                        tb.Fields("BatchNo") = Val(sBatchNo)
                        tb.Fields("BatchStatus") = sBatchStatus
                        tb.Fields("JournalAmount") = sJournalAmount
                        tb.Fields("AmountDeposited") = sAmountDeposited
                        tb.Fields("TransNo") = Val(sTransNo)
                        UpdateRs tb
                    End If
                End If
            End If
        End Select
NextLine:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    Err.Clear
End Sub
Public Sub ImportRegisterOfReceipts(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim sReceiptNo As String
    Dim sSeqNo As String
    Dim sIndicator As String
    Dim sReceiptStatus As String
    Dim sReceiptDate As String
    Dim sUserId As String
    Dim sName As String
    Dim sDescription As String
    Dim sPaymentMethod As String
    Dim sBatchNo As String
    Dim sAmount As String
    fName = FileToken(strFile, "fo")
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    StatusMessage frmObj, "Importing register of receipts..."
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        sReceiptDate = Trim$(Mid$(strLine, 39, 12))
        If IsDate(sReceiptDate) = True Then
            sReceiptNo = Trim$(Mid$(strLine, 1, 8))
            sSeqNo = Trim$(Mid$(strLine, 10, 9))
            sIndicator = Trim$(Mid$(strLine, 20, 6))
            sReceiptStatus = Trim$(Mid$(strLine, 32, 6))
            sUserId = Trim$(Mid$(fData(rsCnt + 1), 40, 12))
            sName = Trim$(Mid$(strLine, 52, 37))
            sDescription = Trim$(Mid$(fData(rsCnt + 1), 53, 37))
            sPaymentMethod = Trim$(Mid$(strLine, 90, 6))
            sBatchNo = Trim$(Mid$(strLine, 97, 11))
            sAmount = ProperAmount(Trim$(Mid$(strLine, 109)))
            Select Case sReceiptNo
            Case "1. DATE", "2. DATE", "DATE FRO"
            Case Else
                Set tb = SeekRs("ReceiptNo", sReceiptNo, "Register Of Receipts")
                If tb.EOF = True Then tb.AddNew
                tb.Fields("ReceiptNo") = sReceiptNo
                tb.Fields("SeqNo") = Val(sSeqNo)
                tb.Fields("Indicator") = sIndicator
                tb.Fields("ReceiptStatus") = sReceiptStatus
                tb.Fields("UserId") = sUserId
                tb.Fields("Name") = sName
                tb.Fields("Description") = sDescription
                tb.Fields("PaymentMethod") = sPaymentMethod
                tb.Fields("BatchNo") = Val(sBatchNo)
                tb.Fields("Amount") = sAmount
                tb.Fields("ReceiptDate") = sReceiptDate
                UpdateRs tb
            End Select
        End If
NextLine:
        DoEvents
    Next
    Set tb = Nothing
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Execute "delete from `Register Of Receipts` where SeqNo = 0;"
    Err.Clear
End Sub
Public Sub Ledger_Reconcile(frmObj As Form, ByVal strItem As String)
    On Error Resume Next
    Dim tbL As New ADODB.Recordset
    Dim tbR As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sTransaction As String
    Dim sPersal As String
    Dim sAmount As String
    Dim oTransaction As String
    Dim oAmount As String
    Dim spTran() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim myTran As String
    Dim repPos As Long
    Dim sPeriod As String
    Dim eDate As String
    Dim sId As String
    Dim xCount As Integer
    fy = ReadRecordToMv("param", "tablename", "fy", "sequence")
    fys = ReadRecordToMv("param", "tablename", "fys", "sequence")
    resp = MyPrompt("Do you want to recompile the reconciliation for the period specified?", "yn", "q", "Confirm Reconciliation: " & fy & "-" & fys)
    CreateTableWithIndexNames UsrName & "ReconcileLedger", "MatchingField,Transactions,Amount", _
    "text         ,memo        ,text", "255          ,            ,255", _
    "MatchingField,Amount", , , , , "MatchingField"
    sPeriod = fys
    eDate = SwapDate(StartEndDate(sPeriod, "e"))
    frmQA.Caption = "eFas - Ledger Reconciliation For " & ProperCase(strItem) & " As At " & sPeriod
    frmQA.lstReport.Tag = "ledgeritems"
    If resp = vbNo Then GoTo Reporting
    Execute "update ledger set Reference = '0' where item = '" & EscIn(strItem) & "' and BasDate <= '" & eDate & "';"
    Set tbL = OpenRs("select transaction,persal,amount from ledger where item = '" & EscIn(strItem) & "' and BasDate <= '" & eDate & "';")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Reconciling ledger account..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbL.Fields("Transaction"))
        sPersal = MyRN(tbL.Fields("Persal"))
        sAmount = ProperCase(MyRN(tbL.Fields("Amount")))
        If Len(sPersal) = 0 Then GoTo NextRecord
        Set tbR = SeekRs("MatchingField", sPersal, UsrName & "ReconcileLedger")
        Select Case tbR.EOF
        Case True
            tbR.AddNew
            tbR.Fields("MatchingField") = sPersal
            tbR.Fields("Transactions") = sTransaction
            tbR.Fields("Amount") = sAmount
            UpdateRs tbR
        Case Else
            oTransaction = MyRN(tbR.Fields("Transactions"))
            oAmount = MyRN(tbR.Fields("Amount"))
            oTransaction = oTransaction & "," & sTransaction
            oAmount = Val(oAmount) + Val(sAmount)
            oAmount = ProperAmount(oAmount)
            tbR.Fields("Transactions") = oTransaction
            tbR.Fields("Amount") = oAmount
            UpdateRs tbR
        End Select
NextRecord:
        DoEvents
        tbL.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tbL.Close
    ' marking matched records, please wait...
    Set tbR = OpenRs(UsrName & "ReconcileLedger")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Marking matched records..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbR.Fields("Transactions"))
        sAmount = ProperAmount(MyRN(tbR.Fields("Amount")))
        spTot = StrParse(spTran, sTransaction, ",")
        Select Case sAmount
        Case "0.00"
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "1"
                    UpdateRs tbL
                End Select
            Next
        Case Else
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "0"
                    UpdateRs tbL
                End Select
            Next
        End Select
        DoEvents
        tbR.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Ledger_ReconcileRemaining frmObj, strItem
Reporting:
    Ledger_ReconcileRemaining frmObj, strItem
    qrySql = "select transaction,EntryType,FuncArea,District,Narration,BasAudit,User,Beneficiary,persal,Reference,debt,SourceDocument,BasDate,Amount from ledger where Item = '" & EscIn(strItem) & "' and EntryType not in ('OB','PO') and BasDate <= '" & eDate & "' and Reference = '0' order by basdate;"
    ViewSQLNew qrySql, frmQA.lstReport, "ID,Func Area,Func Area No,District,Narration,Audit,User,Beneficiary,MF1,Reference,Debt,Source,Date,Amount", , , , , , , , , "Amount"
    LstViewSumColumns frmObj.lstReport, True, "Amount"
    ResetFilter frmObj.lstReport, frmObj.lstValues, frmObj.cboBox, frmObj.chkRemove
    StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " record(s) listed."
    Err.Clear
End Sub
Public Sub Ledger_ReconcileRemaining(frmObj As Form, ByVal strItem As String)
    On Error Resume Next
    Dim tbL As New ADODB.Recordset
    Dim tbR As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sTransaction As String
    Dim sPersal As String
    Dim sAmount As String
    Dim oTransaction As String
    Dim oAmount As String
    Dim spTran() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim myTran As String
    Dim repPos As Long
    Dim sPeriod As String
    Dim eDate As String
    Dim sId As String
    Dim xCount As Integer
    Dim oAmounts As String
    fy = ReadRecordToMv("param", "tablename", "fy", "sequence")
    fys = ReadRecordToMv("param", "tablename", "fys", "sequence")
    CreateTableWithIndexNames UsrName & "ReconcileLedger", _
    "MatchingField,Transactions,Amounts,Amount", _
    "text         ,memo        ,memo   ,text", _
    "255          ,            ,       ,255", _
    "MatchingField,Amount", , , , , "MatchingField"
    sPeriod = fys
    eDate = SwapDate(StartEndDate(sPeriod, "e"))
    Set tbL = OpenRs("select transaction,persal,amount from ledger where item = '" & EscIn(strItem) & "' and BasDate <= '" & eDate & "' and reference = '0';")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Reconciling ledger account..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbL.Fields("Transaction"))
        sPersal = MyRN(tbL.Fields("Persal"))
        sAmount = ProperCase(MyRN(tbL.Fields("Amount")))
        If Len(sPersal) = 0 Then GoTo NextRecord
        Set tbR = SeekRs("MatchingField", sPersal, UsrName & "ReconcileLedger")
        Select Case tbR.EOF
        Case True
            tbR.AddNew
            tbR.Fields("MatchingField") = sPersal
            tbR.Fields("Transactions") = sTransaction
            tbR.Fields("Amount") = sAmount
            tbR.Fields("Amounts") = sAmount
            UpdateRs tbR
        Case Else
            oTransaction = MyRN(tbR.Fields("Transactions"))
            oAmount = MyRN(tbR.Fields("Amount"))
            oAmounts = MyRN(tbR.Fields("Amounts"))
            oTransaction = oTransaction & "," & sTransaction
            oAmounts = oAmounts & "," & sAmount
            oAmount = CDbl(oAmount) + CDbl(sAmount)
            oAmount = ProperAmount(oAmount)
            tbR.Fields("Transactions") = oTransaction
            tbR.Fields("Amount") = oAmount
            tbR.Fields("Amounts") = oAmounts
            UpdateRs tbR
        End Select
NextRecord:
        DoEvents
        tbL.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tbL.Close
    ' marking matched records, please wait...
    Set tbR = OpenRs(UsrName & "ReconcileLedger")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Marking matched records..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbR.Fields("Transactions"))
        sAmount = ProperAmount(MyRN(tbR.Fields("Amount")))
        spTot = StrParse(spTran, sTransaction, ",")
        Select Case sAmount
        Case "0.00"
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "1"
                    UpdateRs tbL
                End Select
            Next
        Case Else
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "0"
                    UpdateRs tbL
                End Select
            Next
        End Select
        DoEvents
        tbR.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Err.Clear
End Sub
Public Sub ImportDebtorsList(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim eDate As String
    Dim sId As String
    Dim sDebtorNo As String
    Dim sDebtNo As String
    Dim sDebtorType As String
    Dim sEntityNo As String
    Dim sLastName As String
    Dim sInitials As String
    Dim sDebtType As String
    Dim sRemainingPeriod As String
    Dim sInstallments As String
    Dim sInterestRate As String
    Dim sRunningInterest As String
    Dim sTotalInterest As String
    Dim sCapitalOutstanding As String
    Dim sBalanceOutStanding As String
    fName = FileToken(strFile, "fo")
    eDate = StartEndDate(fName, "e")
    If IsDate(eDate) = False Then
        Call MyPrompt("The file name for the Register Of Debtors should be in yyyymm format." & vbCr & "This file " & strFile & " does not meet that requirement.", "o", "e", "File Error")
    Err.Clear
        Exit Sub
    End If
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    'db.Execute "delete from `Register Of Debtors` where Yyyymm = " & fName
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing debtors list, please be patient..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        fPart = StringPart(strLine, 1, " ")
        Select Case fPart
        Case "PERSAL", "IDNO", "SUPPLI", "PASPRT", "SAPVEN", "DEPTNO"
            sDebtorNo = Trim$(Mid$(strLine, 61, 6))
            sDebtNo = Trim$(Mid$(strLine, 75, 7))
            sDebtorType = Trim$(Mid$(strLine, 1, 13))
            sEntityNo = Trim$(Mid$(strLine, 15, 15))
            sLastName = Trim$(Mid$(strLine, 31, 16))
            sInitials = Trim$(Mid$(strLine, 58, 2))
            sDebtType = Trim$(Mid$(strLine, 68, 6))
            sRemainingPeriod = Trim$(Mid$(strLine, 83, 3))
            sInstallments = DebtAmount(Trim$(Mid$(strLine, 87, 15)))
            sInterestRate = Trim$(Mid$(strLine, 103, 5))
            sRunningInterest = DebtAmount(Trim$(Mid$(strLine, 109, 12)))
            sTotalInterest = DebtAmount(Trim$(Mid$(strLine, 122, 14)))
            sCapitalOutstanding = DebtAmount(Trim$(Mid$(strLine, 137, 15)))
            sBalanceOutStanding = DebtAmount(Trim$(Mid$(strLine, 153, 15)))
            sId = sDebtorNo & "-" & sDebtNo
            Set tb = SeekRs("ID", sId, "Register Of Debtors")
            If tb.EOF = True Then tb.AddNew
            tb.Fields("ID") = sId
            tb.Fields("DebtorNo") = Val(sDebtorNo)
            tb.Fields("DebtNo") = Val(sDebtNo)
            tb.Fields("DebtorType") = sDebtorType
            tb.Fields("EntityNo") = sEntityNo
            tb.Fields("LastName") = sLastName
            tb.Fields("Initials") = sInitials
            tb.Fields("DebtType") = sDebtType
            tb.Fields("RemainingPeriod") = Val(sRemainingPeriod)
            tb.Fields("Installments") = sInstallments
            tb.Fields("InterestRate") = sInterestRate
            tb.Fields("RunningInterest") = sRunningInterest
            tb.Fields("TotalInterest") = sTotalInterest
            tb.Fields("CapitalOutstanding") = sCapitalOutstanding
            tb.Fields("BalanceOutStanding") = sBalanceOutStanding
            tb.Fields("Yyyymm") = Val(fName)
            tb.Fields("AsAtDate") = eDate
            UpdateRs tb
        End Select
NextLine:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    Err.Clear
End Sub
Public Sub ImportDebtAgeAnalysis(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim eDate As String
    Dim sId As String
    Dim sDebtorNo As String
    Dim sDebtNo As String
    Dim sDebtorName As String
    Dim sLastCredited As String
    Dim sAge03 As String
    Dim sAge36 As String
    Dim sAge612 As String
    Dim sAge12 As String
    Dim sAge23 As String
    Dim sAge3 As String
    Dim sYyyymm As String
    Dim sAsAtDate As String
    Dim sAmount As String
    Dim sDebt As String
    Dim lBlank As String
    Dim sSortCriteria As String
    Dim strPart As String
    Dim strFil As String
    fName = FileToken(strFile, "fo")
    eDate = StartEndDate(fName, "e")
    If IsDate(eDate) = False Then
        Call MyPrompt("The file name for the debt age analysis file should be in yyyymm format." & vbCr & "This file " & strFile & " does not meet that requirement.", "o", "e", "File Error")
    Err.Clear
        Exit Sub
    End If
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    Execute "delete from `Debt Age Analysis` where Yyyymm = " & fName
    StatusMessage frmObj, "Importing debt age analysis..."
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        strPart = StringPart(strLine, 1, " ")
        If IsNumeric(strPart) = False Then GoTo NextLine
        If InStr(strPart, ".") > 0 Then GoTo NextLine
        strFil = String$(7 - Len(strPart), " ")
        strLine = strFil & strLine
        sDebtNo = Trim$(Left$(strLine, 7))
        sDebtorNo = Trim$(Mid$(strLine, 10, 9))
        sDebtorName = Trim$(Mid$(strLine, 21, 32))
        sLastCredited = Trim$(Mid$(strLine, 56, 10))
        sAge03 = DebtAmount(Trim$(Mid$(strLine, 67, 16)))
        sAge36 = DebtAmount(Trim$(Mid$(strLine, 84, 16)))
        sAge612 = DebtAmount(Trim$(Mid$(strLine, 101, 16)))
        sAge12 = DebtAmount(Trim$(Mid$(strLine, 118, 16)))
        sAge23 = DebtAmount(Trim$(Mid$(strLine, 135, 16)))
        sAge3 = DebtAmount(Trim$(Mid$(strLine, 152)))
        sAmount = Val(ProperAmount(sAge03)) + Val(ProperAmount(sAge36)) + Val(ProperAmount(sAge612)) + Val(ProperAmount(sAge12)) + Val(ProperAmount(sAge23)) + Val(ProperAmount(sAge3))
        sAmount = ProperAmount(sAmount)
        If IsDate(sLastCredited) = True Then
            sDebt = sDebtorNo & "-" & sDebtNo
            sId = sDebtorNo & "-" & sDebtNo & "-" & fName
            Set tb = SeekRs("ID", sId, "Debt Age Analysis")
            If tb.EOF = True Then tb.AddNew
            tb.Fields("ID") = sId
            tb.Fields("DebtorNo") = Val(sDebtorNo)
            tb.Fields("DebtNo") = Val(sDebtNo)
            tb.Fields("DebtorName") = sDebtorName
            tb.Fields("LastCredited") = sLastCredited
            tb.Fields("Age03") = sAge03
            tb.Fields("Age36") = sAge36
            tb.Fields("Age612") = sAge612
            tb.Fields("Age12") = sAge12
            tb.Fields("Age23") = sAge23
            tb.Fields("Age3") = sAge3
            tb.Fields("Yyyymm") = Val(fName)
            tb.Fields("AsAtDate") = eDate
            tb.Fields("Amount") = sAmount
            tb.Fields("Debt") = sDebt
            UpdateRs tb
        End If
NextLine:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    Err.Clear
End Sub
Public Sub ImportLeaveEntitlement(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim spLine() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim eDate As String
    Dim sId As String
    Dim sPersalNo As String
    Dim sAppointNo As String
    Dim sSurname As String
    Dim sInitials As String
    Dim sCurrentCycleAccrual As String
    Dim sCurrentCycleProRataCredit As String
    Dim sCurrentCycleRandValue As String
    Dim sPreviousCycleLeaveCredit As String
    Dim sPreviousCycleRandValue As String
    Dim sCappedLeaveCredit As String
    Dim sCappedLeaveRandValue As String
    Dim sAmount As String
    Dim sYyyymm As String
    Dim sAsAtDate As String
    fName = FileToken(strFile, "fo")
    eDate = StartEndDate(fName, "e")
    If IsDate(eDate) = False Then
        Call MyPrompt("The file name for the leave entitlement should be in yyyymm format." & vbCr & "This file " & strFile & " does not meet that requirement.", "o", "e", "File Error")
    Err.Clear
        Exit Sub
    End If
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing leave entitlement..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strLine = Trim$(fData(rsCnt))
        If Len(strLine) = 0 Then GoTo NextLine
        Call StrParse(spLine, strLine, ",")
        ArrayTrimItems spLine
        ReDim Preserve spLine(11)
        sPersalNo = spLine(1)
        sAppointNo = spLine(2)
        sSurname = spLine(3)
        sInitials = spLine(4)
        sCurrentCycleAccrual = Replace$(spLine(5), "..", ".")
        sCurrentCycleProRataCredit = Replace$(spLine(6), "..", ".")
        sCurrentCycleRandValue = ProperAmount(Replace$(spLine(7), "..", "."))
        sPreviousCycleLeaveCredit = Replace$(spLine(8), "..", ".")
        sPreviousCycleRandValue = ProperAmount(Replace$(spLine(9), "..", "."))
        sCappedLeaveCredit = Replace$(spLine(10), "..", ".")
        sCappedLeaveRandValue = ProperAmount(Replace$(spLine(11), "..", "."))
        If Val(sCappedLeaveCredit) < 0 Then
            If Left$(sCappedLeaveRandValue, 1) <> "-" Then sCappedLeaveRandValue = "-" & sCappedLeaveRandValue
        End If
        sAmount = Val(sCurrentCycleRandValue) + Val(sPreviousCycleRandValue) + Val(sCappedLeaveRandValue)
        sAmount = ProperAmount(sAmount)
        Select Case Left$(sPersalNo, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            sId = sPersalNo & "*" & fName
            Set tb = SeekRs("ID", sId, "Leave Entitlement")
            If tb.EOF = True Then tb.AddNew
            tb.Fields("ID") = sId
            tb.Fields("PersalNo") = Val(sPersalNo)
            tb.Fields("AppointNo") = sAppointNo
            tb.Fields("surname") = sSurname
            tb.Fields("Initials") = sInitials
            tb.Fields("CurrentCycleAccrual") = ProperAmount(Val(sCurrentCycleAccrual))
            tb.Fields("CurrentCycleProRataCredit") = ProperAmount(Val(sCurrentCycleProRataCredit))
            tb.Fields("CurrentCycleRandValue") = ProperAmount(Val(sCurrentCycleRandValue))
            tb.Fields("PreviousCycleLeaveCredit") = ProperAmount(Val(sPreviousCycleLeaveCredit))
            tb.Fields("PreviousCycleRandValue") = ProperAmount(Val(sPreviousCycleRandValue))
            tb.Fields("CappedLeaveCredit") = ProperAmount(Val(sCappedLeaveCredit))
            tb.Fields("CappedLeaveRandValue") = ProperAmount(Val(sCappedLeaveRandValue))
            tb.Fields("Amount") = sAmount
            tb.Fields("Yyyymm") = Val(fName)
            tb.Fields("AsAtDate") = eDate
            UpdateRs tb
        End Select
NextLine:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    Err.Clear
End Sub
Public Sub ImportRegisterOfPaymentsFile(frmObj As Form, strFile As String)
    On Error Resume Next
    Dim strLine As String
    Dim fData() As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim fPart As String
    Dim fName As String
    Dim tb As New ADODB.Recordset
    Dim myRecs As String
    Dim sPaymentNo As String
    Dim sMicrNo As String
    Dim sDisbursementNo As String
    Dim sCapturedDate As String
    Dim sSourceDocNo As String
    Dim sStatus As String
    Dim sPaymentMethod As String
    Dim sPayeeName As String
    Dim sDuplicateIndicator As String
    Dim sType As String
    Dim sAmount As String
    Dim tFile As String
    Dim intFree As Long
    StatusMessage frmObj, "Reading " & FileToken(strFile, "fo")
    fName = FileToken(strFile, "fo")
    tFile = App.Path & "\Preprocess.txt"
    If FileExists(tFile) = True Then Kill tFile
    strLine = FileData(strFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    Execute "delete from `Register Of Payments` where yyyymm = " & fName
    Set tb = OpenRs("Register Of Payments", , 1)
    intFree = FreeFile
    StatusMessage frmObj, "Preprocessing " & FileToken(strFile, "fo")
    Open tFile For Output Access Write As #intFree
        For rsCnt = 1 To rsTot
            strLine = Trim$(fData(rsCnt))
            If Len(strLine) = 0 Then GoTo NextLine
            fPart = MvField(strLine, 1, " ")
            Select Case IsNumeric(fPart)
            Case True
                Select Case fPart
                Case "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.", "11."
                Case Else
                    If InStr(fPart, ",") = 0 Then
                        Print #intFree, strLine
                    End If
                End Select
            Case False
                Select Case fPart
                Case "Y", "N"
                    strLine = String(32, " ") & strLine
                    Print #intFree, strLine
                Case "SUNDRY", "INVOIC"
                    strLine = String(32, " ") & "N          " & strLine
                    Print #intFree, strLine
                End Select
            End Select
NextLine:
        Next
    Close #intFree
    strLine = FileData(tFile)
    rsTot = StrParse(fData, strLine, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importinng " & FileToken(strFile, "fo")
    For rsCnt = 1 To rsTot
        strLine = fData(rsCnt)
        frmObj.progBar.Value = rsCnt
        fPart = MvField(strLine, 1, " ")
        Select Case Len(fPart)
        Case 9
            sPaymentNo = Trim$(Mid$(strLine, 1, 9))
            sMicrNo = Trim$(Mid$(strLine, 11, 10))
            sDisbursementNo = Trim$(Mid$(strLine, 22, 10))
            sCapturedDate = Trim$(Mid$(strLine, 33, 10))
            sSourceDocNo = Trim$(Mid$(strLine, 44, 32))
            sStatus = Trim$(Mid$(strLine, 77, 6))
            sPaymentMethod = Trim$(Mid$(strLine, 84, 6))
            sPayeeName = Trim$(Mid$(strLine, 91, 32))
            sDuplicateIndicator = Trim$(Mid$(fData(rsCnt + 1), 33, 10))
            sType = Trim$(Mid$(fData(rsCnt + 1), 44, 32))
            sAmount = ProperAmount(Trim$(Mid$(fData(rsCnt + 1), 77)))
            tb.AddNew
            tb.Fields("PaymentNo") = Val(sPaymentNo)
            tb.Fields("MicrNo") = Val(sMicrNo)
            tb.Fields("DisbursementNo") = Val(sDisbursementNo)
            tb.Fields("CapturedDate") = sCapturedDate
            tb.Fields("SourceDocNo") = sSourceDocNo
            tb.Fields("Status") = sStatus
            tb.Fields("PaymentMethod") = sPaymentMethod
            tb.Fields("PayeeName") = sPayeeName
            tb.Fields("DuplicateIndicator") = sDuplicateIndicator
            tb.Fields("Type") = sType
            tb.Fields("Amount") = sAmount
            tb.Fields("Yyyymm") = Val(Format$(sCapturedDate, "yyyymm"))
            UpdateRs tb
        End Select
        DoEvents
    Next
    tb.Close
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Err.Clear
End Sub
Public Sub Ledger_Reconcile_Selected(frmObj As Form, ByVal strItem As String)
    On Error Resume Next
    Dim tbL As New ADODB.Recordset
    Dim tbR As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sTransaction As String
    Dim sPersal As String
    Dim sAmount As String
    Dim oTransaction As String
    Dim oAmount As String
    Dim spTran() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim myTran As String
    Dim repPos As Long
    Dim sPeriod As String
    Dim eDate As String
    Dim sId As String
    Dim xCount As Integer
    fy = ReadRecordToMv("param", "tablename", "fy", "sequence")
    fys = ReadRecordToMv("param", "tablename", "fys", "sequence")
    CreateTableWithIndexNames UsrName & "ReconcileLedger", "MatchingField,Transactions,Amount", "text         ,memo        ,text", "255          ,            ,255", "MatchingField,Amount", , , , , "MatchingField"
    sPeriod = fys
    eDate = SwapDate(StartEndDate(sPeriod, "e"))
    frmQA.Caption = "eFas - Ledger Reconciliation For " & ProperCase(strItem) & " As At " & sPeriod
    frmQA.lstReport.Tag = "ledgeritems"
    Execute "update ledger set Reference = '0' where item = '" & EscIn(strItem) & "' and BasDate <= '" & eDate & "';"
    Set tbL = OpenRs("select transaction,persal,amount from ledger where item = '" & EscIn(strItem) & "' and BasDate <= '" & eDate & "';")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbL.Fields("Transaction"))
        sPersal = MyRN(tbL.Fields("Persal"))
        sAmount = ProperCase(MyRN(tbL.Fields("Amount")))
        If Len(sPersal) = 0 Then GoTo NextRecord
        Set tbR = SeekRs("MatchingField", sPersal, UsrName & "ReconcileLedger")
        Select Case tbR.EOF
        Case True
            tbR.AddNew
            tbR.Fields("MatchingField") = sPersal
            tbR.Fields("Transactions") = sTransaction
            tbR.Fields("Amount") = sAmount
            UpdateRs tbR
        Case Else
            oTransaction = MyRN(tbR.Fields("Transactions"))
            oAmount = MyRN(tbR.Fields("Amount"))
            oTransaction = oTransaction & "," & sTransaction
            oAmount = Val(oAmount) + Val(sAmount)
            oAmount = ProperAmount(oAmount)
            tbR.Fields("Transactions") = oTransaction
            tbR.Fields("Amount") = oAmount
            UpdateRs tbR
        End Select
NextRecord:
        DoEvents
        tbL.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tbL.Close
    ' marking matched records, please wait...
    Set tbR = OpenRs(UsrName & "ReconcileLedger")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sTransaction = MyRN(tbR.Fields("Transactions"))
        sAmount = MyRN(tbR.Fields("Amount"))
        spTot = StrParse(spTran, sTransaction, ",")
        Select Case sAmount
        Case "0.00"
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "1"
                    UpdateRs tbL
                End Select
            Next
        Case Else
            For spCnt = 1 To spTot
                Set tbL = SeekRs("transaction", spTran(spCnt), "ledger")
                Select Case tbL.EOF
                Case False
                    tbL.Fields("Reference") = "0"
                    UpdateRs tbL
                End Select
            Next
        End Select
        DoEvents
        tbR.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    qrySql = "select transaction,EntryType,FuncArea,District,Responsibility,Objective,Item,project,Narration,BasAudit,User,Beneficiary,persal,debt,SourceDocument,BasDate,yyyymm,Amount from ledger where " & "Item = '" & EscIn(strItem) & "' and EntryType not in ('OB','PO') and BasDate <= '" & eDate & "' and Reference = '0' order by basdate;"
    ViewSQLNew qrySql, frmQA.lstReport, LedgerHeading
    xCount = 0
    StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " record(s) listed."
Redo:
    ClearTable UsrName & "ReconcileLedger"
    ' try and re reconcile any matching transactions left out
    rsTot = frmObj.lstReport.ListItems.Count
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "trying to match remaining records..."
    frmObj.progBar.Value = rsTot
    For rsCnt = rsTot To 1 Step -1
        frmObj.progBar.Value = rsCnt
        spTran = LstViewGetRow(frmQA.lstReport, rsCnt)
        sTransaction = spTran(1)
        sPersal = spTran(13)
        sAmount = ProperAmount(spTran(18))
        sId = sPersal & "*" & ProperAmount(Val(sAmount) * (0 - 1))
        Set tbR = SeekRs("matchingfield", sId, UsrName & "ReconcileLedger")
        Select Case tbR.EOF
        Case True
            tbR.AddNew
            tbR.Fields("MatchingField") = sPersal & "*" & sAmount
            tbR.Fields("Transactions") = sTransaction
            UpdateRs tbR
        Case Else
            oTransaction = MyRN(tbR.Fields("Transactions"))
            frmQA.lstReport.ListItems.Remove rsCnt
            repPos = LstViewFindItem(frmQA.lstReport, oTransaction, search_Text, search_Whole)
            If repPos > 0 Then
                frmQA.lstReport.ListItems.Remove repPos
            End If
            DeleteRs tbR
            'tbR.Delete
        End Select
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " record(s) listed."
    xCount = xCount + 1
    If xCount = 6 Then
        GoTo Finalize
    Else
        GoTo Redo
    End If
Finalize:
    LstViewSumColumns frmQA.lstReport, True, "Amount"
    LstViewAutoResize frmQA.lstReport
    frmQA.lstReport.Refresh
    ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " record(s) listed."
    DeleteTables UsrName & "ReconcileLedger"
    Err.Clear
End Sub
Sub DependencyCount()
    On Error Resume Next
    Dim iU As Long
    Dim ocxList As New Collection
    Dim dllList As New Collection
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sDep As String
    Dim batFile As String
    Dim nDep As String
    batFile = App.Path & "\sysupdate.bat"
    If FileExists(batFile) = True Then Kill batFile
    Set ocxList = MyFilesCollection("z:\eFas\LiveUpdate", "*.ocx")
    Set dllList = MyFilesCollection("z:\eFas\LiveUpdate", "*.dll")
    SaveReg "dependencies", ocxList.Count + dllList.Count, "account", App.Title
    iU = 0
    rsTot = ocxList.Count
    For rsCnt = 1 To rsTot
        sDep = ocxList(rsCnt)
        nDep = GetSysDir & "\" & FileToken(sDep, "f")
        FileUpdate batFile, GetSysDir & "\REGSVR32.EXE" & " " & nDep, "a"
        iU = iU + 1
        SaveReg "dependency" & iU, sDep, "account", App.Title
    Next
    rsTot = dllList.Count
    For rsCnt = 1 To rsTot
        sDep = dllList(rsCnt)
        nDep = GetSysDir & "\" & FileToken(sDep, "f")
        FileUpdate batFile, GetSysDir & "\REGSVR32.EXE" & " " & nDep, "a"
        iU = iU + 1
        SaveReg "dependency" & iU, sDep, "account", App.Title
    Next
    Err.Clear
End Sub
Sub Read_GsscDetailed(ByVal strFile As String)
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim tbLI As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim lenFile As Long
    Dim lngFile As Long
    Dim strLine As String
    Dim spLine() As String
    Dim sdd As String
    Dim smm As String
    Dim syy As String
    Dim sDate As String
    Dim sFo As String
    Dim eDate As String
    sFo = FileToken(strFile, "fo")
    sDate = StartEndDate(sFo, "s")
    eDate = StartEndDate(sFo, "e")
    If IsDate(sDate) = False Or IsDate(eDate) = False Then
        Call MyPrompt("The file name should be in YYYYMM format, please rename the file to meet this requirement.", "o", "e", "File Name Error")
    Err.Clear
        Exit Sub
    End If
    lngFile = FreeFile
    Open strFile For Input Access Read As #lngFile
        rsTot = LOF(lngFile)
        ClearTable "LedgerImport"
        Set tb = OpenRs("LedgerImport")
        rsCnt = 0
        Do Until EOF(lngFile)
            Line Input #lngFile, strLine
            rsCnt = rsCnt + Len(strLine)
            strLine = Trim$(strLine)
            If Len(strLine) = 0 Then GoTo NextRecord
            strLine = Replace$(strLine, Quote, "")
            Call StrParse(spLine, strLine, ",")
            ArrayTrimItems spLine
            sDate = MvField(spLine(16), 1, " ")
            sdd = MvField(sDate, 1, "/")
            smm = MvField(sDate, 2, "/")
            syy = MvField(sDate, 3, "/")
            sdd = StrFormat(sdd, "R%2")
            smm = StrFormat(smm, "R%2")
            sDate = sdd & "/" & smm & "/" & syy
            tb.AddNew
            tb.Fields("RespCode") = MvField(spLine(5), 1, ".")
            tb.Fields("Responsibility") = spLine(6)
            tb.Fields("objCode") = MvField(spLine(7), 1, ".")
            tb.Fields("Objective") = spLine(8)
            tb.Fields("Project") = MvField(spLine(9), 1, ".")
            tb.Fields("ItemCode") = MvField(spLine(10), 1, ".")
            tb.Fields("Item") = spLine(11)
            tb.Fields("entrytype") = spLine(2)
            tb.Fields("funcArea") = Val(spLine(3))
            tb.Fields("narration") = spLine(12)
            tb.Fields("Reference") = spLine(13)
            tb.Fields("BasAudit") = Val(spLine(14))
            tb.Fields("User") = spLine(15)
            tb.Fields("Yyyymm") = Val(spLine(4))
            tb.Fields("Amount") = spLine(17)
            tb.Fields("basdate") = sDate
            tb.Fields("district") = Replace$(StringPart(spLine(6), 1, " "), ":", "")
            If MyRN(tb.Fields("district")) = "CONTROL" Then tb.Fields("district") = "HO"
            UpdateRs tb
NextRecord:
        Loop
        tb.Close
    Close #lngFile
    Set tb = Nothing
    Err.Clear
End Sub
Sub Section21PartialUnder(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - Under Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - Under Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' let's remove all full function 20 statuses and full function 21 statuses
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                Select Case LCase$(section20)
                Case "no,no,no,no", "yes,yes,yes,yes"
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            AmountsBasedOnFunctions frmQA.lstReport
            RemoveZeroTransfers
            VarianceLessGreat "g"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "g"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21PartialUnderByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    Dim dCnt As Long
    Dim dTot As Long
    Dim sDistricts() As String
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - Under Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - Under Expenditure)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "g"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Public Sub VarianceLessGreat(Optional ByVal L_or_G As String = "g")
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    rsTot = frmQA.lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
        Select Case LCase$(L_or_G)
        Case "l"
            If ProperAmount(spLines(19)) = "0.00" Then
                frmQA.lstReport.ListItems.Remove rsCnt
            ElseIf Val(ProperAmount(spLines(19))) > 0 Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Case "g"
            If ProperAmount(spLines(19)) = "0.00" Then
                frmQA.lstReport.ListItems.Remove rsCnt
            ElseIf Val(ProperAmount(spLines(19))) < 0 Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Case "n"
            If ProperAmount(spLines(18)) = "0.00" Then
            Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        End Select
    Next
    Err.Clear
End Sub
Public Sub RemoveZeroTransfers()
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    rsTot = frmQA.lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
        If ProperAmount(spLines(18)) = "0.00" Then
            frmQA.lstReport.ListItems.Remove rsCnt
        End If
    Next
    Err.Clear
End Sub
Public Sub RemoveZeroEverything()
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    rsTot = frmQA.lstReport.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
        If ProperAmount(spLines(12)) = "0.00" And ProperAmount(spLines(18)) = "0.00" And ProperAmount(spLines(19)) = "0.00" Then
            frmQA.lstReport.ListItems.Remove rsCnt
        End If
    Next
    Err.Clear
End Sub
Sub Section21PartialOver(strSelection As String, Optional Revised As Boolean = True)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - Over Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - Over Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' let's remove all full function 20 statuses and full function 21 statuses
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                Select Case LCase$(section20)
                Case "no,no,no,no", "yes,yes,yes,yes"
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            AmountsBasedOnFunctions frmQA.lstReport
            VarianceLessGreat "l"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        VarianceLessGreat "l"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21PartialOverByDistrict(strSelection As String, Optional Revised As Boolean = True)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistricts() As String
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - Over Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - Over Expenditure)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        VarianceLessGreat "l"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section21PartialNone(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - No Transfers)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - No Transfers)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' let's remove all full function 20 statuses and full function 21 statuses
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                Select Case LCase$(section20)
                Case "no,no,no,no", "yes,yes,yes,yes"
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            AmountsBasedOnFunctions frmQA.lstReport
            VarianceLessGreat "n"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        VarianceLessGreat "n"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21PartialNoneByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistricts() As String
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial No Transfers)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial No Transfers)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        VarianceLessGreat "n"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section21PartialIncomplete(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial - Incomplete)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial - Incomplete)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' let's remove all full function 20 statuses and full function 21 statuses
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                Select Case LCase$(section20)
                Case "no,no,no,no", "yes,yes,yes,yes"
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            ' correct amounts based on functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = 1 To rsTot
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If spLines(1) = "Totals" Then GoTo NextLine
                ' if maintenance function is zero, then amount should be zero etc
                If spLines(5) = "No" Then spLines(9) = "0.00"
                If spLines(6) = "No" Then spLines(10) = "0.00"
                If spLines(7) = "No" Then spLines(11) = "0.00"
                ' recalculate total budget to be transferred
                spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
                spLines(12) = MakeMoney(spLines(12))
                ' calculate the variance between budgeted transfers and actual transfers
                spLines(19) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(18)))
                spLines(19) = MakeMoney(spLines(19))
                'update report
                Call LstViewUpdate(spLines, frmQA.lstReport, CStr(rsCnt))
                frmQA.lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
            Next
            ' let's remove all actual transfers = 0
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If Val(ProperAmount(spLines(18))) < Val(ProperAmount(spLines(12))) Then
                Else
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' let's remove all actual transfers = 0
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If ProperAmount(spLines(18)) = "0.00" Then
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            'VarianceLessGreat "n"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        ' correct amounts based on functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = 1 To rsTot
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If spLines(1) = "Totals" Then GoTo NextLine1
            ' if maintenance function is zero, then amount should be zero etc
            If spLines(5) = "No" Then spLines(9) = "0.00"
            If spLines(6) = "No" Then spLines(10) = "0.00"
            If spLines(7) = "No" Then spLines(11) = "0.00"
            ' recalculate total budget to be transferred
            spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
            spLines(12) = MakeMoney(spLines(12))
            ' calculate the variance between budgeted transfers and actual transfers
            spLines(19) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(18)))
            spLines(19) = MakeMoney(spLines(19))
            'update report
            Call LstViewUpdate(spLines, frmQA.lstReport, CStr(rsCnt))
            frmQA.lstReport.ListItems(rsCnt).EnsureVisible
NextLine1:
        Next
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If Val(ProperAmount(spLines(18))) < Val(ProperAmount(spLines(12))) Then
            Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If ProperAmount(spLines(18)) = "0.00" Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        'VarianceLessGreat "n"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20Full(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            AmountsBasedOnFunctions frmQA.lstReport
            RemoveZeroEverything
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroEverything
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20FullByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistricts() As String
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroEverything
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
    Next
    Err.Clear
End Sub
Sub Section21Partial(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' let's remove all full function 20 statuses and full function 21 statuses
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                Select Case LCase$(section20)
                Case "no,no,no,no", "yes,yes,yes,yes"
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            AmountsBasedOnFunctions frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20Over(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Over Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Over Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            ' correct amounts based on functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = 1 To rsTot
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If spLines(1) = "Totals" Then GoTo NextLine
                ' if maintenance function is zero, then amount should be zero etc
                If spLines(5) = "No" Then spLines(9) = "0.00"
                If spLines(6) = "No" Then spLines(10) = "0.00"
                If spLines(7) = "No" Then spLines(11) = "0.00"
                ' recalculate total budget to be transferred
                spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
                spLines(12) = MakeMoney(spLines(12))
                ' calculate the variance between budgeted transfers and actual transfers
                spLines(19) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(18)))
                spLines(19) = MakeMoney(spLines(19))
                'update report
                Call LstViewUpdate(spLines, frmQA.lstReport, CStr(rsCnt))
                frmQA.lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
            Next
            VarianceLessGreat "l"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        ' correct amounts based on functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = 1 To rsTot
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If spLines(1) = "Totals" Then GoTo NextLine1
            ' if maintenance function is zero, then amount should be zero etc
            If spLines(5) = "No" Then spLines(9) = "0.00"
            If spLines(6) = "No" Then spLines(10) = "0.00"
            If spLines(7) = "No" Then spLines(11) = "0.00"
            ' recalculate total budget to be transferred
            spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
            spLines(12) = MakeMoney(spLines(12))
            ' calculate the variance between budgeted transfers and actual transfers
            spLines(19) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(18)))
            spLines(19) = MakeMoney(spLines(19))
            'update report
            Call LstViewUpdate(spLines, frmQA.lstReport, CStr(rsCnt))
            frmQA.lstReport.ListItems(rsCnt).EnsureVisible
NextLine1:
        Next
        VarianceLessGreat "l"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub AnalyzeTB(frmObj As Form, ByVal strPeriod As String, Optional ByVal LevelId As String = "RECEIPTS", Optional ByVal Level As Integer = 1, Optional AmountLessThanZero As Boolean = True)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim tbTB As New ADODB.Recordset
    Dim tbI As New ADODB.Recordset
    Dim sParent As String
    Dim sitem As String
    Dim sAmount As String
    Dim sPath As String
    Dim spLine(1 To 3) As String
    LstViewMakeHeadings frmObj.lstReport, "Economic| Classification,Account,Amount"
    frmObj.lstReport.View = lvwReport
    frmObj.lstReport.Checkboxes = False
    frmObj.lstReport.GridLines = True
    frmObj.lstReport.FullRowSelect = True
    Set tbTB = OpenRs("select * from `trial balances` where Period = " & strPeriod)
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sParent = MyRN(tbTB.Fields("Parent"))
        sitem = MyRN(tbTB.Fields("Item"))
        sAmount = ProperAmount(MyRN(tbTB.Fields("Amount")))
        Set tbI = SeekRs("PostingLevelKey", sitem & "\Y", "items")
        Select Case tbI.EOF
        Case False
            sPath = MyRN(tbI.Fields("Path"))
            sPath = MvField(sPath, Level, "\")
            Select Case UCase$(sPath)
            Case UCase$(LevelId)
                If AmountLessThanZero = True Then
                    If Val(sAmount) > 0 Then
                        spLine(1) = sParent
                        spLine(2) = sitem
                        spLine(3) = sAmount
                        Call LstViewUpdate(spLine, frmObj.lstReport, "")
                    End If
                ElseIf AmountLessThanZero = False Then
                    If Val(sAmount) < 0 Then
                        spLine(1) = sParent
                        spLine(2) = sitem
                        spLine(3) = sAmount
                        Call LstViewUpdate(spLine, frmObj.lstReport, "")
                    End If
                End If
            End Select
        Case Else
            spLine(1) = "Error:" & sParent
            spLine(2) = sitem
            spLine(3) = sAmount
            Call LstViewUpdate(spLine, frmObj.lstReport, "")
        End Select
        DoEvents
        tbTB.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tbTB.Close
    LstViewSumColumns frmObj.lstReport, True, "Amount"
    ResetFilter frmObj.lstReport, frmObj.lstValue, frmObj.cboField, frmObj.chkRemove
    Set tbTB = Nothing
    Set tbI = Nothing
    Err.Clear
End Sub
Sub AccountNumberUsage(frmObj As Form)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim tb As New ADODB.Recordset
    Dim sBranchCode As String
    Dim sBankAccNo As String
    Dim sSurname As String
    Dim sPath As String
    Dim anu As New Collection
    Dim spRec() As String
    Dim spLine(1 To 2) As String
    Dim rssCnt As Long
    Dim rssTot As Long
    frmQA.Caption = "eFas - Bank Account Number Usage - Report 1"
    Set tb = OpenRs("Payment Advice")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBranchCode = Trim$(MyRN(tb.Fields("branchCode")))
        sBankAccNo = Trim$(MyRN(tb.Fields("BankAccNo")))
        sSurname = Trim$(MyRN(tb.Fields("surname")))
        sBranchCode = Iconv(sBranchCode)
        sBankAccNo = Iconv(sBankAccNo)
        sPath = Trim$(sBankAccNo & " " & sBranchCode & "\" & sSurname)
        If Len(StringPart(sPath, 1, "\")) > 0 Then
            If Len(StringPart(sPath, 2, "\")) > 0 Then
                anu.Add sPath, sPath
            End If
        End If
        DoEvents
        tb.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tb.Close
    CreateTableWithIndexNames UsrName & "BankUsage", "Account,Supplier,Cnt", "Text,Memo,long", "255,,", "Account,Cnt", , , , , "Account"
    rsTot = anu.Count
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBankAccNo = anu(rsCnt)
        sBranchCode = Trim$(MvField(sBankAccNo, 1, "\"))
        sSurname = Trim$(MvField(sBankAccNo, 2, "\"))
        If Len(sBranchCode) = 0 Then GoTo NextItem
        Set tb = SeekRs("Account", sBranchCode, UsrName & "BankUsage")
        Select Case tb.EOF
        Case True
            tb.AddNew
            tb.Fields("Account") = sBranchCode
            tb.Fields("Supplier") = sSurname
            tb.Fields("cnt") = 1
            UpdateRs tb
        Case Else
            tb.Fields("Supplier") = MyRN(tb.Fields("Supplier")) & VM & sSurname
            tb.Fields("cnt") = MvCount(MyRN(tb.Fields("Supplier")), VM)
            UpdateRs tb
        End Select
NextItem:
        DoEvents
    Next
    StatusMessage frmObj, "Loading affected records..."
    ProgBarClose frmObj.progBar
    Execute "delete from `" & UsrName & "BankUsage` where cnt = 1;"
    LstViewMakeHeadings frmQA.lstReport, "Bank Account,Supplier"
    frmQA.lstReport.View = lvwReport
    frmQA.lstReport.FullRowSelect = True
    frmQA.lstReport.Tag = "query"
    Set tb = OpenRs(UsrName & "BankUsage")
    rsTot = AffectedRecords
    For rsCnt = 1 To rsTot
        sBranchCode = MyRN(tb.Fields("Account"))
        sSurname = MyRN(tb.Fields("Supplier"))
        Call StrParse(spRec, sSurname, VM)
        rssTot = UBound(spRec)
        For rssCnt = 1 To rssTot
            spLine(1) = sBranchCode
            spLine(2) = spRec(rssCnt)
            Call LstViewUpdate(spLine, frmQA.lstReport, "")
        Next
        tb.MoveNext
    Next
    tb.Close
    Set anu = Nothing
    Set tb = Nothing
    StatusMessage frmObj
    LstViewAutoResize frmQA.lstReport
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True
    DeleteTables UsrName & "BankUsage"
    Err.Clear
End Sub
Sub AccountNumberUsage_EBT(frmObj As Form)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim tb As New ADODB.Recordset
    Dim sBranchCode As String
    Dim sBankAccNo As String
    Dim sSurname As String
    Dim sPath As String
    Dim anu As New Collection
    Dim spRec() As String
    Dim spLine(1 To 3) As String
    Dim rssCnt As Long
    Dim rssTot As Long
    Dim sCount As String
    frmQA.Caption = "eFas - Ebt Payments Bank Account Number Usage - Report"
    Set tb = OpenRs("Ebt Payments")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBranchCode = Trim$(MyRN(tb.Fields("BranchName")))
        sBankAccNo = Trim$(MyRN(tb.Fields("AccountNo")))
        sSurname = Trim$(MyRN(tb.Fields("Beneficiary")))
        sPath = Trim$(sBankAccNo & " " & sBranchCode & "\" & sSurname)
        If Len(StringPart(sPath, 1, "\")) > 0 Then
            If Len(StringPart(sPath, 2, "\")) > 0 Then
                anu.Add sPath, sPath
            End If
        End If
        DoEvents
        tb.MoveNext
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    tb.Close
    CreateTableWithIndexNames UsrName & "BankUsage", "Account,Supplier,Cnt", "Text,Memo,long", "255,,", "Account,Cnt", , , , , "Account"
    rsTot = anu.Count
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Compiling bank account usage, please be patient..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBankAccNo = anu(rsCnt)
        sBranchCode = Trim$(MvField(sBankAccNo, 1, "\"))
        sSurname = Trim$(MvField(sBankAccNo, 2, "\"))
        If Len(sBranchCode) = 0 Then GoTo NextItem
        Set tb = SeekRs("Account", sBranchCode, UsrName & "BankUsage")
        Select Case tb.EOF
        Case True
            tb.AddNew
            tb.Fields("Account") = sBranchCode
            tb.Fields("Supplier") = sSurname
            tb.Fields("cnt") = 1
            UpdateRs tb
        Case Else
            tb.Fields("Supplier") = MyRN(tb.Fields("Supplier")) & VM & sSurname
            tb.Fields("cnt") = MvCount(MyRN(tb.Fields("Supplier")), VM)
            UpdateRs tb
        End Select
NextItem:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj, "Loading affected records..."
    Execute "delete from " & UsrName & "BankUsage where cnt = 1;"
    LstViewMakeHeadings frmQA.lstReport, "Bank Account,Beneficiary,Count"
    frmQA.lstReport.View = lvwReport
    frmQA.lstReport.FullRowSelect = True
    frmQA.lstReport.Tag = "query"
    Set tb = OpenRs(UsrName & "BankUsage")
    rsTot = AffectedRecords
    For rsCnt = 1 To rsTot
        sBranchCode = MyRN(tb.Fields("Account"))
        sSurname = MyRN(tb.Fields("Supplier"))
        sCount = MyRN(tb.Fields("cnt"))
        Call StrParse(spRec, sSurname, VM)
        rssTot = UBound(spRec)
        For rssCnt = 1 To rssTot
            spLine(1) = sBranchCode
            spLine(2) = spRec(rssCnt)
            spLine(3) = sCount
            Call LstViewUpdate(spLine, frmQA.lstReport, "")
        Next
        tb.MoveNext
    Next
    tb.Close
    Set anu = Nothing
    Set tb = Nothing
    StatusMessage frmObj
    frmQA.lstReport.ColumnHeaders(3).Alignment = lvwColumnRight
    LstViewAutoResize frmQA.lstReport
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True
    DeleteTables UsrName & "BankUsage"
    Err.Clear
End Sub
Sub AccountNumberUsageReport2(frmObj As Form)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim tb As New ADODB.Recordset
    Dim sBranchCode As String
    Dim sBankAccNo As String
    Dim sSurname As String
    Dim sPath As String
    Dim anu As New Collection
    Dim spRec() As String
    Dim spLine(1 To 2) As String
    Dim rssCnt As Long
    Dim rssTot As Long
    frmQA.Caption = "eFas - Bank Account Number Usage - Report 2"
    Set tb = OpenRs("Payment Advice")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBranchCode = Trim$(MyRN(tb.Fields("branchCode")))
        sBankAccNo = Trim$(MyRN(tb.Fields("BankAccNo")))
        sSurname = Trim$(MyRN(tb.Fields("surname")))
        sBranchCode = Iconv(sBranchCode)
        sBankAccNo = Iconv(sBankAccNo)
        sPath = Trim$(sSurname & "\" & sBankAccNo & " " & sBranchCode)
        If Len(StringPart(sPath, 1, "\")) > 0 Then
            If Len(StringPart(sPath, 2, "\")) > 0 Then
                anu.Add sPath, sPath
            End If
        End If
        DoEvents
        tb.MoveNext
    Next
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    tb.Close
    CreateTableWithIndexNames UsrName & "BankUsage", "Account,Supplier,Cnt", "Text,Memo,long", "255,,", "Account,Cnt", , , , , "Account"
    rsTot = anu.Count
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Compiling bank account usage..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sBankAccNo = anu(rsCnt)
        sBranchCode = Trim$(MvField(sBankAccNo, 1, "\"))
        sSurname = Trim$(MvField(sBankAccNo, 2, "\"))
        If Len(sBranchCode) = 0 Then GoTo NextItem
        Set tb = SeekRs("Account", sBranchCode, UsrName & "BankUsage")
        Select Case tb.EOF
        Case True
            tb.AddNew
            tb.Fields("Account") = sBranchCode
            tb.Fields("Supplier") = sSurname
            tb.Fields("cnt") = 1
            UpdateRs tb
        Case Else
            tb.Fields("Supplier") = MyRN(tb.Fields("Supplier")) & VM & sSurname
            tb.Fields("cnt") = MvCount(MyRN(tb.Fields("Supplier")), VM)
            UpdateRs tb
        End Select
NextItem:
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj, "Loading affected records..."
    Execute "delete from `" & UsrName & "BankUsage` where cnt = 1;"
    LstViewMakeHeadings frmQA.lstReport, "Supplier,Bank Account"
    frmQA.lstReport.View = lvwReport
    frmQA.lstReport.FullRowSelect = True
    frmQA.lstReport.Tag = "query"
    Set tb = OpenRs(UsrName & "BankUsage")
    rsTot = AffectedRecords
    For rsCnt = 1 To rsTot
        sBranchCode = MyRN(tb.Fields("Account"))
        sSurname = MyRN(tb.Fields("Supplier"))
        Call StrParse(spRec, sSurname, VM)
        rssTot = UBound(spRec)
        For rssCnt = 1 To rssTot
            spLine(1) = sBranchCode
            spLine(2) = spRec(rssCnt)
            Call LstViewUpdate(spLine, frmQA.lstReport, "")
        Next
        tb.MoveNext
    Next
    tb.Close
    Set anu = Nothing
    Set tb = Nothing
    StatusMessage frmObj
    LstViewAutoResize frmQA.lstReport
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True
    DeleteTables UsrName & "BankUsage"
    Err.Clear
End Sub
Sub DetailedExceptionsReport(frmObj As Form)
    On Error Resume Next
    Dim tbS As New ADODB.Recordset
    Dim tbT As New ADODB.Recordset
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim sDecision As String
    Dim sExceptions As String
    Dim spLine() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim sSerial As String
    Dim sFindings As String
    Dim myLine As String
    Dim sOffice As String
    Dim sStatus As String
    Dim sFullName As String
    Dim spRecord(1 To 7) As String
    Dim sSurname As String
    Dim sCompiler As String
    Dim sChecker As String
    Dim sAuthorizer As String
    Dim compilerPos As Long
    Dim checkerPos As Long
    Dim authoPos As Long
    frmQA.Caption = "eFas - Detailed Exceptions Report As At " & Format$(Now(), "dd/mm/yyyy hh:mm")
    LstViewMakeHeadings frmQA.lstReport, "Payment No,Office,Supplier,Exception,Compiler,Checker,Authorizer"
    Set tbS = OpenRs("select Exceptions,Serial,findings,Office,Status,FullName,Surname from `Payment Advice` where decision = 'EXCEPTION' OR decision = 'FAILED';")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Compiling detailed exceptions report..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sExceptions = MyRN(tbS.Fields("Exceptions"))
        sSerial = MyRN(tbS.Fields("Serial"))
        sFindings = MyRN(tbS.Fields("Findings"))
        sOffice = MyRN(tbS.Fields("Office"))
        sStatus = MyRN(tbS.Fields("Status"))
        sFullName = MyRN(tbS.Fields("FullName"))
        sSurname = MyRN(tbS.Fields("surname"))
        'CompilerýCheckerýAuthorizer
        compilerPos = MvSearch(sStatus, "Compiler", VM)
        checkerPos = MvSearch(sStatus, "Checker", VM)
        authoPos = MvSearch(sStatus, "Authorizer", VM)
        sCompiler = MvField(sFullName, compilerPos, VM)
        sChecker = MvField(sFullName, checkerPos, VM)
        sAuthorizer = MvField(sFullName, authoPos, VM)
        If Len(sExceptions) > 0 Then
            spTot = StrParse(spLine, sExceptions, VM)
            For spCnt = 1 To spTot
                myLine = Trim$(spLine(spCnt))
                If Len(myLine) = 0 Then GoTo NextLine
                If LCase$(Left$(myLine, 14)) <> "recommendation" Then
                    spRecord(1) = sSerial
                    spRecord(2) = sOffice
                    spRecord(3) = sSurname
                    spRecord(4) = myLine
                    spRecord(5) = ProperCase(sCompiler)
                    spRecord(6) = ProperCase(sChecker)
                    spRecord(7) = ProperCase(sAuthorizer)
                    Call LstViewUpdate(spRecord, frmQA.lstReport, "")
                End If
NextLine:
            Next
        End If
        If Len(sFindings) = 0 Then GoTo NextRecord
        spTot = StrParse(spLine, sFindings, VM)
        For spCnt = 1 To spTot
            myLine = Trim$(spLine(spCnt))
            If Len(myLine) = 0 Then GoTo NextLine1
            spRecord(1) = sSerial
            spRecord(2) = sOffice
            spRecord(3) = sSurname
            spRecord(4) = myLine
            spRecord(5) = ProperCase(sCompiler)
            spRecord(6) = ProperCase(sChecker)
            spRecord(7) = ProperCase(sAuthorizer)
            Call LstViewUpdate(spRecord, frmQA.lstReport, "")
NextLine1:
        Next
NextRecord:
        DoEvents
        tbS.MoveNext
    Next
    tbS.Close
    Set tbS = Nothing
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Err.Clear
End Sub
Sub BeneficiaryFromPersal(frmObj As Form)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim tbE As New ADODB.Recordset
    Dim tbL As New ADODB.Recordset
    Dim spLine() As String
    Dim sPersal As String
    Dim sName As String
    rsTot = frmQA.lstReport.ListItems.Count
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        spLine = LstViewGetRow(frmQA.lstReport, rsCnt)
        If spLine(1) = "Totals" Then GoTo NextRecord
        Set tbL = SeekRs("transaction", spLine(1), "ledger")
        Select Case tbL.EOF
        Case False
            sPersal = MyRN(tbL.Fields("Persal"))
            If Len(sPersal) = 0 Then GoTo NextRecord
            Set tbE = SeekRs("persal", sPersal, "employees")
            Select Case tbE.EOF
            Case False
                sName = MyRN(tbE.Fields("fullname"))
            Case Else
                sName = ""
            End Select
            tbE.Close
            tbL.Fields("Beneficiary") = sName
            UpdateRs tbL
            spLine(12) = UCase$(sName)
            Call LstViewUpdate(spLine, frmQA.lstReport, CStr(rsCnt))
        End Select
NextRecord:
        DoEvents
    Next
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Set tbL = Nothing
    Set tbE = Nothing
    Err.Clear
End Sub
Sub Section21Full(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
            AmountsBasedOnFunctions frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance"
        AmountsBasedOnFunctions frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21FullGS(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Goods And Services)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Goods And Services)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctionsGS frmQA.lstReport
            LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsGS frmQA.lstReport
        LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21FullGSByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim sDistrict As String
    Dim sDistricts() As String
    Dim dTot As Long
    Dim dCnt As Long
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Goods And Services)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Goods And Services)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsGS frmQA.lstReport
        LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section20FullGS(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full - Goods And Services)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full - Goods And Services)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctionsGS frmQA.lstReport
            LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsGS frmQA.lstReport
        LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20Both(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full - Goods And Services And Transfers)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full - Goods And Services And Transfers)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' and CurrentTransfers > 0 order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' and CurrentTransfers > 0 order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' and CurrentTransfers > 0 order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' and CurrentTransfers > 0 order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 1
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub RevisedFunctions(strSelection As String, Optional ByDistrict As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim oldF As String
    Dim newF As String
    Dim spDistricts() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim myDistricts As String
    Dim sDistrict As String
    mySelection = "Schools With Revised Function For Year " & strSelection
    If ByDistrict = True Then
        mySelection = "Schools With Revised Function For Year " & strSelection & " By District"
    End If
    If ByDistrict = True Then
        qrySql = "select distinct district from `schools budget` where year = " & strSelection
        myDistricts = DistinctColumnString(qrySql, "district", ";", True)
        spTot = StrParse(spDistricts, myDistricts, ";")
        For spCnt = 1 To spTot
            sDistrict = spDistricts(spCnt)
            frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Revised |Budget| Maintenance,Revised| Budget| LSM,Revised| Budget| Services,Revised| Total| Budget", , , , , True, , , , "Maintenance,LSM,Services,Total"
            ' remove same functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                oldF = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
                oldF = LCase$(oldF)
                newF = LCase$(newF)
                If oldF = newF Then
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            AmountsBasedOnFunctionsRevised frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Revised |Budget| Maintenance", "Revised| Budget| LSM", "Revised| Budget| Services", "Revised| Total| Budget"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Next
    Else
        frmQA.Caption = mySelection
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Revised |Budget| Maintenance,Revised| Budget| LSM,Revised| Budget| Services,Revised| Total| Budget", , , , , True, , , , "Maintenance,LSM,Services,Total"
        ' remove same functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            oldF = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
            oldF = LCase$(oldF)
            newF = LCase$(newF)
            If oldF = newF Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        AmountsBasedOnFunctionsRevised frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Revised |Budget| Maintenance", "Revised| Budget| LSM", "Revised| Budget| Services", "Revised| Total| Budget"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20NoGS(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full - No Goods And Services)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full - No Goods And Services)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctionsGS frmQA.lstReport
            LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 0
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsGS frmQA.lstReport
        LstViewFilterNew frmQA, frmQA.lstReport, "Actual| Goods & Services", "0.00", 0
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20NoGSTransfers(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full - Transfers But No Goods And Services)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full - Transfers But No Goods And Services)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' and CurrentTransfers > 0 and CurrentGoodsAndServices = 0 order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' and CurrentTransfers > 0 and CurrentGoodsAndServices = 0 order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' and CurrentTransfers > 0 and CurrentGoodsAndServices = 0 order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' and CurrentTransfers > 0 and CurrentGoodsAndServices = 0 order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section20Nothing(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 - Full - No Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 20 - Full - No Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            LstViewFilterNew frmQA, frmQA.lstReport, "Total| Expenditure", "0.00", 0
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' and Section21 = 'No' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and RFuncA = 'No' and RFuncC = 'No' and RFuncD = 'No' and RSection21 = 'No' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        LstViewFilterNew frmQA, frmQA.lstReport, "Total| Expenditure", "0.00", 0
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21Over(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Over Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Over Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctions frmQA.lstReport
            RemoveZeroTransfers
            VarianceLessGreat "l"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "l"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21OverByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistrict As String
    Dim sDistricts() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Over Expenditure)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Over Expenditure)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        StatusMessage frmQA, "Processing " & dCnt & " of " & dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "l"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section21Under(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    If Revised = False Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Under Expenditure)"
    Else
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Under Expenditure)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctions frmQA.lstReport
            RemoveZeroTransfers
            VarianceLessGreat "g"
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "g"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21UnderByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistrict As String
    Dim sDistricts() As String
    If Revised = False Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full Under Expenditure)"
    Else
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 Full - Under Expenditure)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroTransfers
        VarianceLessGreat "g"
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section21None(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - No Transfers)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - No Transfers)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            ' let's remove all actual transfers = 0
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                Select Case ProperAmount(spLines(18))
                Case "0.00"
                Case Else
                    frmQA.lstReport.ListItems.Remove rsCnt
                End Select
            Next
            AmountsBasedOnFunctions frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            Select Case ProperAmount(spLines(18))
            Case "0.00"
            Case Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21NoneByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim sDistricts() As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full No Transfers)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full No Transfers)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            Select Case ProperAmount(spLines(18))
            Case "0.00"
            Case Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub Section21Incomplete(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full - Incomplete)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full - Incomplete)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            ' let's remove all actual transfers = 0
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If Val(ProperAmount(spLines(18))) < Val(ProperAmount(spLines(12))) Then
                Else
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' let's remove all actual transfers = 0
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                If ProperAmount(spLines(18)) = "0.00" Then
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If Val(ProperAmount(spLines(18))) < Val(ProperAmount(spLines(12))) Then
            Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' let's remove all actual transfers = 0
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            If ProperAmount(spLines(18)) = "0.00" Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Public Function ImportRevisedSchoolBudget(frmObj As Form, StrSource As String) As Boolean
    On Error Resume Next
    Dim dbs As DAO.Database
    Dim tbS As DAO.Recordset
    Dim tbT As New ADODB.Recordset
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim sSection21 As String
    Dim sFuncA As String
    Dim sFuncC As String
    Dim sFuncD As String
    Dim sYear As String
    Dim sId As String
    Dim sTotal As String
    If DaoTableExists(StrSource, "schoolbudget") = False Then
        ImportRevisedSchoolBudget = False
    Err.Clear
        Exit Function
    End If
    Set dbs = DAO.OpenDatabase(StrSource)
    Set tbS = dbs.OpenRecordset("schoolbudget")
    sYear = FileToken(StrSource, "fo")
    rsTot = tbS.RecordCount
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        sId = RN(tbS.Fields("ref")) & "-" & sYear
        Set tbT = SeekRs("ID", sId, "Schools Budget")
        If tbT.EOF = True Then tbT.AddNew
        sSection21 = RN(tbS.Fields("section21"))
        sFuncA = Trim$(RN(tbS.Fields("FuncA")))
        sFuncC = Trim$(RN(tbS.Fields("FuncC")))
        sFuncD = Trim$(RN(tbS.Fields("FuncD")))
        If sSection21 = "0" Then sSection21 = "No"
        If sSection21 = "-1" Then sSection21 = "Yes"
        If sSection21 = "False" Then sSection21 = "No"
        If sSection21 = "True" Then sSection21 = "Yes"
        If sFuncA = "0" Then sFuncA = "No"
        If sFuncA = "-1" Then sFuncA = "Yes"
        If sFuncA = "False" Then sFuncA = "No"
        If sFuncA = "True" Then sFuncA = "Yes"
        If sFuncC = "0" Then sFuncC = "No"
        If sFuncC = "-1" Then sFuncC = "Yes"
        If sFuncC = "False" Then sFuncC = "No"
        If sFuncC = "True" Then sFuncC = "Yes"
        If sFuncD = "0" Then sFuncD = "No"
        If sFuncD = "-1" Then sFuncD = "Yes"
        If sFuncD = "False" Then sFuncD = "No"
        If sFuncD = "True" Then sFuncD = "Yes"
        tbT.Fields("ID") = sId
        tbT.Fields("Year") = sYear
        tbT.Fields("EMIS") = RN(tbS.Fields("ref"))
        tbT.Fields("district") = RN(tbS.Fields("DIS"))
        tbT.Fields("schoolname") = ProperCase(RN(tbS.Fields("School Name")))
        tbT.Fields("rsection21") = sSection21
        tbT.Fields("rFuncA") = sFuncA
        tbT.Fields("rFuncC") = sFuncC
        tbT.Fields("rFuncD") = sFuncD
        UpdateRs tbT
        DoEvents
        tbS.MoveNext
    Next
    tbS.Close
    dbs.Close
    Set tbT = Nothing
    Set tbS = Nothing
    Set dbs = Nothing
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    ImportRevisedSchoolBudget = True
    Err.Clear
End Function
Sub SchoolsOnIds(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Schools With IDS And LSM Transfers For Year " & strSelection
    If Revised = True Then
        mySelection = "Schools With IDS And LSM Transfers For Year " & strSelection & " (Revised)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Budget| Rtt,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Rtt,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctionsIDS frmQA.lstReport
            RttAndLsm frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance", "Budget| Rtt", "Actual| Rtt"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Budget| Rtt,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Rtt,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsIDS frmQA.lstReport
        RttAndLsm frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance", "Budget| Rtt", "Actual| Rtt"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
    Err.Clear
End Sub
Sub SchoolsOnIdsRep2(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Schools With IDS Transfers For Year " & strSelection
    If Revised = True Then
        mySelection = "Schools With IDS Transfers For Year " & strSelection & " (Revised)"
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Budget| Rtt,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Rtt,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctionsIDS frmQA.lstReport
            Rtt frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance", "Budget| Rtt", "Actual| Rtt"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Budget| Rtt,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Rtt,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsIDS frmQA.lstReport
        Rtt frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance", "Budget| Rtt", "Actual| Rtt"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
    Err.Clear
End Sub
Sub SchoolsOnIdsByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim sDistrict As String
    Dim sDistricts() As String
    Dim dCnt As Long
    Dim dTot As Long
    mySelection = "Schools With IDS Transfers For Year " & strSelection
    If Revised = True Then
        mySelection = "Schools With IDS Transfers For Year " & strSelection & " (Revised)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " For " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Rtt,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentRtt,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Budget| Rtt,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Rtt,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctionsIDS frmQA.lstReport
        Rtt frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance", "Budget| Rtt", "Actual| Rtt"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
    Next
    Err.Clear
End Sub
Sub Section20To21(strSelection As String)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim oldF As String
    Dim newF As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 20 To Section 21)"
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' order by district,schoolname;"
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            ' remove same functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                oldF = StringToMv(",", spLines(5), spLines(6), spLines(7))
                newF = StringToMv(",", spLines(9), spLines(10), spLines(11))
                oldF = LCase$(oldF)
                newF = LCase$(newF)
                If oldF = newF Then
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' remove irrelevant functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                newF = StringToMv(",", spLines(9), spLines(10), spLines(11))
                newF = LCase$(newF)
                If newF = "yes,yes,yes" Then
                Else
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'No' and FuncC = 'No' and FuncD = 'No' order by district,schoolname;"
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' remove same functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            oldF = StringToMv(",", spLines(5), spLines(6), spLines(7))
            newF = StringToMv(",", spLines(9), spLines(10), spLines(11))
            oldF = LCase$(oldF)
            newF = LCase$(newF)
            If oldF = newF Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' remove irrelevant functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            newF = StringToMv(",", spLines(9), spLines(10), spLines(11))
            newF = LCase$(newF)
            If newF = "yes,yes,yes" Then
            Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21To20(strSelection As String)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim oldF As String
    Dim newF As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 To Section 20)"
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            ' remove same functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                oldF = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
                newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
                oldF = LCase$(oldF)
                newF = LCase$(newF)
                If oldF = newF Then
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' remove irrelevant functions
            rsTot = frmQA.lstReport.ListItems.Count
            For rsCnt = rsTot To 1 Step -1
                spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
                newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
                newF = LCase$(newF)
                If newF = "no,no,no,no" Then
                Else
                    frmQA.lstReport.ListItems.Remove rsCnt
                End If
            Next
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
            LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentGoodsAndServices,CurrentTotal,Variance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' order by district,schoolname;"
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Rev| Func A,Rev| Func C,Rev| Func D,Rev| Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Actual| Goods & Services,Total| Expenditure,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' remove same functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            oldF = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
            oldF = LCase$(oldF)
            newF = LCase$(newF)
            If oldF = newF Then
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' remove irrelevant functions
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            newF = StringToMv(",", spLines(9), spLines(10), spLines(11), spLines(12))
            newF = LCase$(newF)
            If newF = "no,no,no,no" Then
            Else
                frmQA.lstReport.ListItems.Remove rsCnt
            End If
        Next
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Expenditure", "Variance", "Actual| Goods & Services"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Sub Section21FullByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistricts() As String
    Dim sDistrict As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Full)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Full)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and FuncA = 'Yes' and FuncC = 'Yes' and FuncD = 'Yes' and Section21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and RFuncA = 'Yes' and RFuncC = 'Yes' and RFuncD = 'Yes' and RSection21 = 'Yes' and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        RemoveZeroEverything
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
    Next
    Err.Clear
End Sub
Sub Section21PartialByDistrict(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    Dim section20 As String
    Dim dTot As Long
    Dim dCnt As Long
    Dim sDistrict As String
    Dim sDistricts() As String
    mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Section 21 - Partial)"
    If Revised = True Then
        mySelection = "Reconciliation Of Transfers To Schools For Year " & strSelection & " (Revised Section 21 - Partial)"
    End If
    sDistrict = DistinctColumnString("select distinct district from `schools budget` where year = " & strSelection, "district", ";")
    sDistrict = MvSort_String(sDistrict, ";")
    dTot = StrParse(sDistricts, sDistrict, ";")
    For dCnt = 1 To dTot
        sDistrict = sDistricts(dCnt)
        frmQA.Caption = mySelection & " " & DistrictFullName(sDistrict)
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities," & "CurrentSchoolSupport,CurrentTransfers,TransfersVariance from `schools budget` where year = " & strSelection & " and district = '" & EscIn(sDistrict) & "' order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget,Actual| Maintenance,Actual| LTSM,Actual| Services," & "Actual| Utilities,Actual| School Support,Total| Transfers,Variance", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        ' let's remove all full function 20 statuses and full function 21 statuses
        rsTot = frmQA.lstReport.ListItems.Count
        For rsCnt = rsTot To 1 Step -1
            spLines = LstViewGetRow(frmQA.lstReport, rsCnt)
            section20 = StringToMv(",", spLines(5), spLines(6), spLines(7), spLines(8))
            Select Case LCase$(section20)
            Case "no,no,no,no", "yes,yes,yes,yes"
                frmQA.lstReport.ListItems.Remove rsCnt
            End Select
        Next
        AmountsBasedOnFunctions frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget", "Actual| Maintenance", "Actual| LTSM", "Actual| Services", "Actual| Utilities"
        LstViewSumColumns frmQA.lstReport, True, "Actual| School Support", "Total| Transfers", "Variance"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        PrintExcel App.Path & "\Reports\" & Province & " " & Department, frmQA.Caption, frmQA.lstReport, , False, True, False, True
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    Next
    Err.Clear
End Sub
Sub SbList(strSelection As String, Optional Revised As Boolean = False)
    On Error Resume Next
    Dim mySelection As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    mySelection = "Approved Budget For Schools For Year " & strSelection
    If Revised = True Then
        mySelection = "Revised Budget For Schools For Year " & strSelection
    End If
    frmQA.Caption = mySelection
    If RecordExists("MyReports", "ID", mySelection) = True Then
        resp = MyPrompt("This report already exists, would you like to refresh it or view current. Click Yes to view current and No to refresh the report.", "yn", "q", "Confirm Report")
        If resp = vbYes Then
            LstViewOpenReport frmQA.lstReport, mySelection, , True
            frmQA.Caption = mySelection
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        Else
            qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            If Revised = True Then
                qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " order by district,schoolname;"
            End If
            ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
            AmountsBasedOnFunctions frmQA.lstReport
            ' sum reports
            LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget"
            ' save report
            LstViewAutoResize frmQA.lstReport
            StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
            LstViewSaveReport frmQA.lstReport, frmQA.Caption
            ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
        End If
    Else
        qrySql = "select id,District,SchoolName,Responsibility,FuncA,FuncC,FuncD,Section21,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        If Revised = True Then
            qrySql = "select id,District,SchoolName,Responsibility,RFuncA,RFuncC,RFuncD,RSection21,Maintenance,LSM,Services,Total from `schools budget` where year = " & strSelection & " order by district,schoolname;"
        End If
        ViewSQLNew qrySql, frmQA.lstReport, "id,District,School Name,Responsibility,Func A,Func C,Func D,Section 21,Budget| Maintenance,Budget| LSM,Budget| Services,Total| Budget", , , , , True, , , , "Maintenance,LSM,Services,Total,CurrentMaintenance,CurrentLTSM,CurrentServices,CurrentUtilities,CurrentSchoolSupport,CurrentTransfers,TransfersVariance,CurrentGoodsAndServices,CurrentTotal,Variance"
        AmountsBasedOnFunctions frmQA.lstReport
        ' sum reports
        LstViewSumColumns frmQA.lstReport, True, "Budget| Maintenance", "Budget| LSM", "Budget| Services", "Total| Budget"
        ' save report
        LstViewAutoResize frmQA.lstReport
        StatusMessage frmQA, frmQA.lstReport.ListItems.Count & " schools listed"
        LstViewSaveReport frmQA.lstReport, frmQA.Caption
        ResetFilter frmQA.lstReport, frmQA.lstValue, frmQA.cboField, frmQA.chkRemove
    End If
    Err.Clear
End Sub
Public Sub AmountsBasedOnFunctionsGS(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = 1 To rsTot
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If LCase$(spLines(5)) = "no" Then spLines(9) = "0.00"
        If LCase$(spLines(6)) = "no" Then spLines(10) = "0.00"
        If LCase$(spLines(7)) = "no" Then spLines(11) = "0.00"
        ' recalculate total budget to be transferred
        spLines(12) = Val(ProperAmount(spLines(9))) + Val(ProperAmount(spLines(10))) + Val(ProperAmount(spLines(11)))
        spLines(12) = MakeMoney(spLines(12))
        ' recalculate actual
        spLines(19) = Val(ProperAmount(spLines(13))) + Val(ProperAmount(spLines(14))) + Val(ProperAmount(spLines(15))) + Val(ProperAmount(spLines(16))) + Val(ProperAmount(spLines(17))) + Val(ProperAmount(spLines(18)))
        spLines(19) = MakeMoney(spLines(19))
        ' calculate the variance between budgeted transfers and actual transfers
        spLines(20) = Val(ProperAmount(spLines(12))) - Val(ProperAmount(spLines(19)))
        spLines(20) = MakeMoney(spLines(20))
        'update report
        Call LstViewUpdate(spLines, lstReport, CStr(rsCnt))
        lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
    Next
    Err.Clear
End Sub
Public Sub AmountsBasedOnFunctionsRevised(lstReport As ListView)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spLines() As String
    ' correct amounts based on functions
    rsTot = lstReport.ListItems.Count
    For rsCnt = 1 To rsTot
        spLines = LstViewGetRow(lstReport, rsCnt)
        If spLines(1) = "Totals" Then GoTo NextLine
        ' if maintenance function is zero, then amount should be zero etc
        If LCase$(spLines(5)) = "no" Then spLines(13) = "0.00"
        If LCase$(spLines(6)) = "no" Then spLines(14) = "0.00"
        If LCase$(spLines(7)) = "no" Then spLines(15) = "0.00"
        If LCase$(spLines(9)) = "no" Then spLines(17) = "0.00"
        If LCase$(spLines(10)) = "no" Then spLines(18) = "0.00"
        If LCase$(spLines(11)) = "no" Then spLines(19) = "0.00"
        ' recalculate total budget to be transferred
        spLines(16) = Val(ProperAmount(spLines(13))) + Val(ProperAmount(spLines(14))) + Val(ProperAmount(spLines(15)))
        spLines(16) = MakeMoney(spLines(16))
        ' recalculate revised budget to be transferred
        spLines(20) = Val(ProperAmount(spLines(17))) + Val(ProperAmount(spLines(18))) + Val(ProperAmount(spLines(19)))
        spLines(20) = MakeMoney(spLines(20))
        'update report
        Call LstViewUpdate(spLines, lstReport, CStr(rsCnt))
        lstReport.ListItems(rsCnt).EnsureVisible
NextLine:
    Next
    Err.Clear
End Sub
Public Sub Bas_ReadAuditTrail(frmObj As Form, ByVal strFile As String, Optional TargetTable As String = "AuditTrail", Optional projectThere As Boolean = True, Optional theOperation As Boolean = True)
    On Error Resume Next
    Dim lngFile As Long
    Dim rsStr As String
    Dim fPart As String
    Dim rsPrev As String
    Dim postDate As String
    Dim auditNum As String
    Dim tranDate As String
    Dim tranNum As String
    Dim tranTyp As String
    Dim tranUsr As String
    Dim itmNumb As String
    Dim itmCode As String
    Dim itmDesc As String
    Dim itmDetl As String
    Dim itmAmnt As String
    Dim itmSign As String
    Dim lineMFD1 As String
    Dim lineMFV1 As String
    Dim lineMFD2 As String
    Dim lineMFV2 As String
    Dim lineMFD3 As String
    Dim lineMFV3 As String
    Dim ObjCode As String
    Dim objDesc As String
    Dim rspCode As String
    Dim rspDesc As String
    Dim fndCode As String
    Dim fndDesc As String
    Dim prjCode As String
    Dim prjDesc As String
    Dim sLineDescription As String
    Dim fileLines() As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim tb As New ADODB.Recordset
    Dim aKey As String
    StatusMessage frmObj, "Reading contents of file " & FileToken(strFile, "fo")
    rsStr = FileData(strFile)
    rsTot = StrParse(fileLines, rsStr, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        rsStr = fileLines(rsCnt)
        fPart = StringPart(rsStr, 1, " ")
        Select Case fPart
        Case "BAS", "RP0043BS", "REPORT", "------------------------", "INSTALLATION", "LOCATION", "USERID:", "SORT", "Sort", "SELECTION", "Selection", "1.", "2.", "3.", "4.", "POST", "DATE", "----------", "LINE", "USERID"
            GoTo NextLine
        End Select
        If IsDate(fPart) = True Then
            postDate = Trim$(Left$(rsStr, 10))
            auditNum = Trim$(Mid$(rsStr, 12, 9))
            tranDate = Trim$(Mid$(rsStr, 22, 11))
            tranNum = Trim$(Mid$(rsStr, 72, 10))    ' func number
            tranTyp = Trim$(Mid$(rsStr, 84, 6))     ' func type
            tranUsr = Trim$(Mid$(rsStr, 91, 12))    ' authorising user
        End If
        ' save previous line in a variable
        Select Case UCase$(fPart)
        Case "O"
            ' objective allocation
            ' ensure the line is the same length as the item line
            rsStr = String$(4, " ") & rsStr
            ObjCode = Trim$(Mid$(rsStr, 13, 9))
            objDesc = Trim$(Mid$(rsStr, 23, 33))
            ' we need to read item data first from the previous line
            ' get the line number of the transaction
            itmNumb = StringPart(rsPrev, 1, " ")
            If Len(itmNumb) = 3 Then
            ElseIf Len(itmNumb) = 2 Then
                rsPrev = " " & rsPrev
            ElseIf Len(itmNumb) = 1 Then
                rsPrev = "  " & rsPrev
            End If
            itmCode = Trim$(Mid$(rsPrev, 7, 15))
            itmDesc = Trim$(Mid$(rsPrev, 23, 33))
            itmDetl = Trim$(Mid$(rsPrev, 57, 33))
            itmAmnt = Trim$(Mid$(rsPrev, 90))
            itmSign = Right$(itmAmnt, 2)
            itmAmnt = Replace$(itmAmnt, "CR", "")
            itmAmnt = Replace$(itmAmnt, "DB", "")
            itmAmnt = Trim$(itmAmnt)
            itmAmnt = ProperAmount(itmAmnt)
            If itmSign = "CR" Then itmAmnt = "-" & itmAmnt
        Case "R"
            ' responsibility allocations
            rsStr = String$(4, " ") & rsStr
            rspCode = Trim$(Mid$(rsStr, 13, 9))
            rspDesc = Trim$(Mid$(rsStr, 23, 33))
        Case "F"
            ' fund allocations
            rsStr = String$(4, " ") & rsStr
            fndCode = Trim$(Mid$(rsStr, 13, 9))
            fndDesc = Trim$(Mid$(rsStr, 23, 33))
            lineMFD1 = Trim$(Mid$(rsStr, 57, 32))
            lineMFV1 = Trim$(Mid$(rsStr, 90))
            If projectThere = False Then
                prjCode = ""
                prjDesc = ""
                lineMFD2 = ""
                lineMFV2 = ""
                auditNum = Val(auditNum)
                tranNum = Val(tranNum)
                itmNumb = Val(itmNumb)
                aKey = auditNum & "." & itmNumb
                If theOperation = False Then
                    DeleteRecord TargetTable, "ID", aKey
                Else
                    Set tb = SeekRs("ID", aKey, TargetTable)
                    If tb.EOF = True Then tb.AddNew
                    tb.Fields("ID") = aKey
                    tb.Fields("Audit") = auditNum
                    tb.Fields("Line") = itmNumb
                    If IsDate(postDate) = True Then tb.Fields("postDate") = postDate
                    If IsDate(tranDate) = True Then tb.Fields("TransactionDate") = tranDate
                    tb.Fields("funcArea") = tranTyp & tranNum
                    tb.Fields("User") = tranUsr
                    tb.Fields("ItemCode") = itmCode
                    tb.Fields("ItemDescription") = itmDesc
                    tb.Fields("ResponsibilityCode") = rspCode
                    tb.Fields("ResponsibilityDescription") = rspDesc
                    tb.Fields("ObjectiveCode") = ObjCode
                    tb.Fields("ObjectiveDesctription") = objDesc
                    tb.Fields("FundCode") = fndCode
                    tb.Fields("FundDescription") = fndDesc
                    tb.Fields("Description") = itmDetl
                    tb.Fields("MatchingFieldDescription1") = lineMFD1
                    tb.Fields("MatchingFieldValue1") = lineMFV1
                    tb.Fields("MatchingFieldDescription2") = lineMFD2
                    tb.Fields("MatchingFieldValue2") = lineMFV2
                    tb.Fields("Amount") = itmAmnt
                    tb.Fields("ProjectDescription") = prjDesc
                    tb.Fields("projectcode") = Val(prjCode)
                    UpdateRs tb
                End If
            End If
        Case "P"
            ' project allocations
            rsStr = String$(4, " ") & rsStr
            prjCode = Trim$(Mid$(rsStr, 13, 9))
            prjDesc = Trim$(Mid$(rsStr, 23, 33))
            lineMFD2 = Trim$(Mid$(rsStr, 57, 32))
            lineMFV2 = Trim$(Mid$(rsStr, 90))
            auditNum = Val(auditNum)
            tranNum = Val(tranNum)
            itmNumb = Val(itmNumb)
            aKey = auditNum & "." & itmNumb
            If projectThere = True Then
                If theOperation = False Then
                    DeleteRecord TargetTable, "ID", aKey
                Else
                    Set tb = SeekRs("ID", aKey, TargetTable)
                    If tb.EOF = True Then tb.AddNew
                    tb.Fields("ID") = aKey
                    tb.Fields("Audit") = auditNum
                    tb.Fields("Line") = itmNumb
                    If IsDate(postDate) = True Then tb.Fields("postDate") = postDate
                    If IsDate(tranDate) = True Then tb.Fields("TransactionDate") = tranDate
                    tb.Fields("funcArea") = tranTyp & tranNum
                    tb.Fields("User") = tranUsr
                    tb.Fields("ItemCode") = itmCode
                    tb.Fields("ItemDescription") = itmDesc
                    tb.Fields("ResponsibilityCode") = rspCode
                    tb.Fields("ResponsibilityDescription") = rspDesc
                    tb.Fields("ObjectiveCode") = ObjCode
                    tb.Fields("ObjectiveDesctription") = objDesc
                    tb.Fields("FundCode") = fndCode
                    tb.Fields("FundDescription") = fndDesc
                    tb.Fields("Description") = itmDetl
                    tb.Fields("MatchingFieldDescription1") = lineMFD1
                    tb.Fields("MatchingFieldValue1") = lineMFV1
                    tb.Fields("MatchingFieldDescription2") = lineMFD2
                    tb.Fields("MatchingFieldValue2") = lineMFV2
                    tb.Fields("Amount") = itmAmnt
                    tb.Fields("ProjectDescription") = prjDesc
                    tb.Fields("projectcode") = Val(prjCode)
                    UpdateRs tb
                End If
            End If
        Case Else
            rsPrev = rsStr
        End Select
NextLine:
        DoEvents
    Next
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Erase fileLines
    Set tb = Nothing
    Err.Clear
End Sub
Sub Bas_UploadDetailedReport(frmObj As Form, ByVal TBase As String, ByVal strFile As String, ByVal depCode As String, Optional BasUnallocated As Boolean = False, Optional MyMonth As String = "")
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim f_Data As String
    Dim spLines() As String
    Dim cLine As String
    Dim sEntryType As String
    Dim sBasTran As String
    Dim sNarration As String
    Dim sBasTran1 As String
    Dim sSysFrom As String
    Dim sBasDate As String
    Dim sAmount As String
    Dim sReference As String
    Dim sDr As String
    Dim sCr As String
    Dim sR As String
    Dim sO As String
    Dim sp As String
    Dim si As String
    Dim sMF As String
    Dim hasDistricts As Boolean
    If ReadRecordToMv("department", "ID", depCode, "HasDistricts") = "0" Then
        hasDistricts = False
    Else
        hasDistricts = True
    End If
    Set tb = OpenRs(TBase, , 1)
    StatusMessage frmQA, "Reading contents of file " & FileToken(strFile, "fo")
    f_Data = FileData(strFile)
    Call StrParse(spLines, f_Data, vbNewLine)
    rsTot = UBound(spLines)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing " & FileToken(strFile, "fo") & ", please wait..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        cLine = Trim$(Replace$(spLines(rsCnt), Quote, ""))
        If Len(cLine) = 0 Then GoTo NextLine
        Select Case Left$(cLine, 2)
        Case "00"
            cLine = Trim$(Mid$(cLine, 3))
        Case "01"
            cLine = Trim$(Mid$(cLine, 3))
        Case "02"
            cLine = Trim$(Mid$(cLine, 3))
        End Select
        'select case
        If UCase$(Left$(cLine, Len("MATCHING FIELD: CLOSING BALANCE"))) = "MATCHING FIELD: CLOSING BALANCE" Then
            sMF = ""
        ElseIf UCase$(Left$(cLine, Len("MATCHING FIELD: OPENING BALANCE"))) = "MATCHING FIELD: OPENING BALANCE" Then
            sMF = Mid$(cLine, 40, 9)
        End If
        If IsUploadable(cLine) = False Then GoTo NextLine
        sEntryType = Trim$(Mid$(cLine, 1, 2))
        Select Case sEntryType
        Case "R"
            sR = Trim$(Mid$(cLine, 9, 60))
        Case "O"
            sO = Trim$(Mid$(cLine, 9, 60))
        Case "P"
            sp = Trim$(Mid$(cLine, 9, 60))
        Case "I"
            si = Trim$(Mid$(cLine, 9, 60))
        Case "GJ", "AP", "DT", "CR", "DR", "TK", "CV", "PO", "BR", "BD", "PS", "DB", "TA"
            sBasTran = Val(Trim$(Mid$(cLine, 4, 10)))
            sNarration = Trim$(Mid$(cLine, 15, 32))
            sBasTran1 = Trim$(Mid$(cLine, 48, 8))
            sSysFrom = Trim$(Mid$(cLine, 57, 12))
            sBasDate = Trim$(Mid$(cLine, 70, 10))
            sDr = Trim$(Mid$(cLine, 81, 17))
            sCr = Trim$(Mid$(cLine, 99))
            sDr = Replace$(sDr, "*", "")
            sCr = Replace$(sCr, "*", "")
            If sDr = "0.00" Then
                sAmount = "-" & sCr
            End If
            If sCr = "0.00" Then
                sAmount = sDr
            End If
            sReference = ExtractNumbers(sNarration)
            If BasUnallocated = True Then
                sReference = ""
                If InStr(1, sNarration, "GGMTFIS") > 0 Then
                    sNarration = Replace$(sNarration, "GGMTFIS", "")
                End If
                sNarration = Trim$(sNarration)
                If Left$(sNarration, 5) = "GGNO:" Then
                    sReference = Trim$(Mid$(sNarration, 6))
                    sReference = Trim$(MvField(sReference, 1, " "))
                End If
                If Left$(sNarration, 11) = "Journal No:" Then
                    sReference = Trim$(Mid$(sNarration, 12))
                End If
                sNarration = Trim$(Replace$(sNarration, "-", " "))
                If Left$(sNarration, Len("JNL NO. ")) = "JNL NO. " Then
                    sReference = MvField(sNarration, 3, " ")
                ElseIf Left$(sNarration, Len("JNL NO.")) = "JNL NO." Then
                    sReference = ExtractNumbers(MvField(sNarration, 2, " "))
                ElseIf Left$(sNarration, Len("JNL NO: ")) = "JNL NO: " Then
                    sReference = MvField(sNarration, 3, " ")
                ElseIf Left$(sNarration, Len("JNL NO:")) = "JNL NO:" Then
                    sReference = ExtractNumbers(MvField(sNarration, 2, " "))
                ElseIf Left$(sNarration, Len("JNL ")) = "JNL " Then
                    sReference = ExtractNumbers(MvField(sNarration, 2, " "))
                ElseIf Left$(sNarration, Len("JNL")) = "JNL" Then
                    sReference = ExtractNumbers(MvField(sNarration, 1, " "))
                End If
            End If
            sAmount = NoCommas(sAmount)
            sR = Replace$(sR, "'", "")
            sO = Replace$(sO, "'", "")
            sp = Replace$(sp, "'", "")
            si = Replace$(si, "'", "")
            If sAmount <> "0.00" Then
                tb.AddNew
                tb.Fields("Responsibility") = sR
                tb.Fields("Objective") = sO
                tb.Fields("Project") = sp
                tb.Fields("Item") = si
                tb.Fields("entrytype") = sEntryType
                tb.Fields("funcArea") = sBasTran
                tb.Fields("narration") = sNarration
                tb.Fields("Reference") = sReference
                tb.Fields("BasAudit") = Val(sBasTran1)
                tb.Fields("User") = sSysFrom
                If IsDate(sBasDate) = True Then tb.Fields("basdate") = sBasDate
                tb.Fields("Amount") = sAmount
                tb.Fields("Persal") = sMF
                'tb.Fields("statusref = ""
                'tb.Fields("StatusDate = Null
                If IsDate(sBasDate) = True Then
                    tb.Fields("Yyyymm") = Format$(sBasDate, "yyyymm")
                Else
                    tb.Fields("Yyyymm") = ""
                End If
                If InStr(1, sR, ":", vbTextCompare) > 0 Then
                    tb.Fields("district") = StringPart(sR, 1, ":")
                Else
                    tb.Fields("district") = Replace$(MvField(sR, 1, " "), ":", "")
                End If
                If MyRN(tb.Fields("district")) = "CONTROL" Then tb.Fields("district") = "HO"
                If hasDistricts = False Then tb.Fields("district") = "HO"
                If Len(MyMonth) > 0 Then
                    If Format$(sBasDate, "yyyymm") = MyMonth Then
                        UpdateRs tb
                    End If
                Else
                    UpdateRs tb
                End If
            End If
        End Select
        DoEvents
NextLine:
    Next
    tb.Close
    Set tb = Nothing
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    Err.Clear
End Sub
Sub Bas_ImportTB(frmObj As Form, ByVal strFile As String)
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim f_Data As String
    Dim spLines() As String
    Dim cLine As String
    Dim sEntryType As String
    Dim sMonth As String
    Dim sKey As String
    Dim sitem As String
    Dim sDebit As String
    Dim sCredit As String
    Dim sAmount As String
    Dim sParent As String
    sMonth = FileToken(strFile, "fo")
    Execute "delete from `Trial Balances` where Period = " & sMonth
    f_Data = FileData(strFile)
    Call StrParse(spLines, f_Data, vbNewLine)
    rsTot = UBound(spLines)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing trial balance, for " & FileToken(strFile, "fo")
    sParent = ""
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        cLine = Trim$(Replace$(spLines(rsCnt), Quote, ""))
        If Len(cLine) = 0 Then GoTo NextLine
        If IsUploadable(cLine) = False Then GoTo NextLine
        sEntryType = Trim$(Mid$(cLine, 1, 2))
        Select Case sEntryType
        Case "I"
            sitem = Trim$(Mid$(cLine, 9, 66))
            sitem = Replace$(sitem, "'", "")
            sDebit = Trim$(Mid$(cLine, 77, 22))
            sCredit = Trim$(Mid$(cLine, 101, 22))
            sDebit = Replace$(sDebit, "*", "")
            sCredit = Replace$(sCredit, "*", "")
            sDebit = ProperAmount(sDebit)
            sCredit = ProperAmount(sCredit)
            If sDebit = "0.00" Then
                sAmount = "-" & sCredit
            End If
            If sCredit = "0.00" Then
                sAmount = sDebit
            End If
            If Mid$(cLine, 1, 5) = "I 003" Then
                sParent = sitem
            End If
            If sAmount <> "0.00" Then
                sKey = sMonth & "*" & sitem
                Set tb = SeekRs("ID", sKey, "Trial Balances")
                If tb.EOF = True Then tb.AddNew
                tb.Fields("ID") = sKey
                tb.Fields("Period") = Val(sMonth)
                tb.Fields("Parent") = sParent
                tb.Fields("Item") = sitem
                tb.Fields("Amount") = sAmount
                tb.Fields("balances") = 0
                tb.Fields("Date") = StartEndDate(sMonth, "e")
                UpdateRs tb
            End If
        End Select
NextLine:
        DoEvents
    Next
    Set tb = Nothing
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Err.Clear
End Sub
Public Sub dbCreateQueryDef(ByVal Ddb As String, ByVal Qryname As String, ByVal qrySql As String)
    On Error Resume Next
    Dim qdf As DAO.QueryDef
    Dim mData As DAO.Database
    Set mData = DAO.OpenDatabase(Ddb)
    Set qdf = mData.CreateQueryDef(Qryname)
    Select Case Err
    Case 3012           ' already exists
        Set qdf = mData.QueryDefs(Qryname)
    End Select
    qdf.sql = qrySql
    mData.Close
    Set mData = Nothing
    Set qdf = Nothing
    Err.Clear
End Sub
Public Sub dbDeleteQueries(ByVal DbName As String, ParamArray items())
    On Error Resume Next
    Dim test As String
    Dim Item As Variant
    Dim db As DAO.Database
    Set db = DAO.OpenDatabase(DbName)
    For Each Item In items
        test = LCase$(CStr(Item))
        test = db.QueryDefs(test).Name
        If Err = 0 Then
            db.QueryDefs.Delete test
        End If
        Err = 0
    Next
    db.Close
    Set db = Nothing
    Err.Clear
End Sub
Sub Bas_CheckAuditTrail(frmObj As Form, ByVal AppPath As String, ByVal myAccount As String, ByVal MyStartDate As String, ByVal MyEndDate As String, Optional ByVal myItem As String = "(All)", Optional ByVal TargetFile As String = "AuditTrail")
    On Error Resume Next
    Dim QuerySQL As String
    Dim sDate As String
    Dim eDate As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim strAT As String
    Dim cMissing As New Collection
    Dim strFile As String
    Dim atRs As New ADODB.Recordset
    Dim maRs As New ADODB.Recordset
    Dim onSource As String
    Dim onTarget As String
    Dim spTarget() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sourcePos As Long
    sDate = MyStartDate
    eDate = MyEndDate
    If myItem = "(All)" Then
        QuerySQL = "select distinct BasAudit from `" & myAccount & "` where basdate >= '" & SwapDate(sDate) & "' and basdate <= '" & SwapDate(eDate) & "' and entrytype not in ('PO','OB')"
    Else
        QuerySQL = "select distinct BasAudit from `" & myAccount & "` where item = '" & EscIn(myItem) & "' and basdate >= '" & SwapDate(sDate) & "' and basdate <= '" & SwapDate(eDate) & "' and entrytype not in ('PO','OB')"
    End If
    StatusMessage frmObj, "Extracting audit trails, please be patient..."
    Set maRs = OpenRs(QuerySQL)
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        strAT = MyRN(maRs.Fields("BasAudit").Value)
        If Len(strAT) > 0 Then
            If RecordExists(TargetFile, "Audit", strAT) = False Then
                cMissing.Add strAT
            End If
        End If
        maRs.MoveNext
        DoEvents
    Next
    StatusMessage frmObj
    ProgBarClose frmObj.progBar
    strAT = MvFromCollection(cMissing, ",")
    strAT = MvRemoveBlanks(strAT, ",")
    maRs.Close
    strFile = FileName_Validate(AppPath & "\" & myAccount & " " & myItem & " MissingAuditTrail 1.txt")
    strAT = SequenceFinder(strAT, ",")
    FileUpdate strFile, strAT, "w"
    Call ViewFile(strFile)
    Set maRs = Nothing
    StatusMessage frmObj, "Selecting source audit trails..."
    QuerySQL = "select distinct BasAudit from `ledger` where item = 'T&S ADVANCE DOM: CA' and entrytype not in ('PO','OB');"
    onSource = DistinctColumnString(QuerySQL, "basaudit", VM)
    DoEvents
    StatusMessage frmObj, "Selecting target audit trails..."
    QuerySQL = "select distinct Audit from `t&s advance dom ca audittrail`;"
    onTarget = DistinctColumnString(QuerySQL, "Audit", VM)
    StatusMessage frmObj, "Checking source trails..."
    Set cMissing = New Collection
    spTot = StrParse(spTarget, onSource, VM)
    ProgBarInit frmObj.progBar, spTot
    For spCnt = 1 To spTot
        frmObj.progBar.Value = spCnt
        onTarget = spTarget(spCnt)
        sourcePos = MvSearch(onTarget, onTarget, VM)
        If sourcePos = 0 Then
            cMissing.Add onTarget
        End If
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    strAT = MvFromCollection(cMissing, ",")
    strAT = MvRemoveBlanks(strAT, ",")
    strFile = FileName_Validate(AppPath & "\" & myAccount & " " & myItem & " MissingAuditTrail 2.txt")
    strAT = SequenceFinder(strAT, ",")
    FileUpdate strFile, strAT, "w"
    Call ViewFile(strFile)
    StatusMessage frmObj
    Err.Clear
End Sub
Public Function DaoTableExists(ByVal Dbase As String, ByVal TbName As String) As Boolean
    On Error Resume Next
    Dim DatCt As Long
    Dim StrDt As String
    Dim zCnt As Long
    Dim db As DAO.Database
    TbName = ProperCase(TbName)
    TbName = Iconv(TbName, "t")
    DaoTableExists = False
    Set db = DAO.OpenDatabase(Dbase)
    With db
        zCnt = .TableDefs.Count - 1
        For DatCt = 0 To zCnt
            StrDt = ProperCase(.TableDefs(DatCt).Name)
            Select Case StrDt
            Case TbName
                DaoTableExists = True
                Exit For
            End Select
        Next
    End With
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Function DebtAmount(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = NoCommas(StrValue)
    Select Case Right$(StrValue, 1)
    Case "-"
        DebtAmount = "-" & Left$(StrValue, Len(StrValue) - 1)
    Case Else
        DebtAmount = StrValue
    End Select
    Err.Clear
End Function
Public Sub MonthlySchedule(frmObject As Form, LstView As ListView, ByVal strTable As String, ByVal StrDateFld As String, ByVal StrGroupFld As String, ByVal StrAmountFld As String, ByVal NewColumnName As String, ByVal sDate As String, ByVal eDate As String, Optional ByVal strSQL As String = "", Optional LastQuery As String = "", Optional SortOrder As String = "", Optional ByVal ConvertDistrict As Boolean = False, Optional RemoveEmptyColumns As Boolean = True, Optional ShowTotalOnly As Boolean = False, Optional MakeUpper As Boolean = False, Optional LimitGroupTo As String)
    On Error Resume Next
    StatusMessage frmObject, "Preparing schedule, please be patient..."
    Dim tbF As New ADODB.Recordset
    Dim tbS As New ADODB.Recordset
    Dim nCnt As Long
    Dim sMonths As String
    Dim lsDate As Long
    Dim leDate As Long
    Dim fName As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim invoiceDate As String
    Dim invoicetotal As String
    Dim xTot() As String
    Dim xFld() As String
    Dim fTot As Long
    Dim fCnt As Long
    Dim invoiceaccount As String
    Dim sql As String
    Dim oldValue As String
    Dim oldMonths As String
    Dim cboCombo As New Collection
    LstView.Sorted = False
    lsDate = Val(DateIconv(SQLDateToNormal(sDate)))
    leDate = Val(DateIconv(SQLDateToNormal(eDate)))
    For nCnt = lsDate To leDate
        sMonths = DateOconv(CStr(nCnt))
        sMonths = Format$(sMonths, "mmm yyyy")
        cboCombo.Add sMonths
    Next
    sMonths = MvFromCollection(cboCombo, ",")
    sMonths = MvRemoveBlanks(sMonths, ",")
    sMonths = MvRemoveDuplicates(sMonths, ",")
    oldMonths = sMonths & ",Totals"
    fTot = MvCount(sMonths, ",")
    ReDim xTot(fTot) As String
    Call StrParse(xFld, sMonths, ",")
    StatusMessage frmObject, "Creating temporal file..."
    CreateTableWithIndexNames UsrName & "MonthlySchedule", StrGroupFld & "," & sMonths & ",Totals", , , StrGroupFld, , , , , StrGroupFld
    sql = "select * from `" & strTable & "` where `" & StrDateFld & "` >= '" & sDate & "' and `" & StrDateFld & "` <= '" & eDate & "'"
    If Len(SortOrder) > 0 Then
        sql = sql & " order by " & SQLQuote(SortOrder)
    Else
        sql = sql & " order by " & SQLQuote(StrDateFld)
    End If
    If Len(strSQL) > 0 Then sql = strSQL
    StatusMessage frmObject, "Selecting transactions to process, please be patient..."
    Set tbF = OpenRs(sql)
    rsTot = AffectedRecords
    ProgBarInit frmObject.progBar, rsTot
    StatusMessage frmObject, "Compiling schedule by " & NewColumnName
    For rsCnt = 1 To rsTot
        frmObject.progBar.Value = rsCnt
        invoiceDate = MyRN(tbF.Fields(StrDateFld))
        invoicetotal = MyRN(tbF.Fields(StrAmountFld))
        invoiceaccount = MyRN(tbF.Fields(StrGroupFld))
        If ConvertDistrict = True Then invoiceaccount = DistrictCode(invoiceaccount) & " - " & invoiceaccount
        invoicetotal = ProperAmount(DotValue(invoicetotal))
        fName = Format$(invoiceDate, "mmm yyyy")
        If Len(invoiceaccount) = 0 Then invoiceaccount = ProperCase("<blank>")
        invoiceaccount = ProperCase(invoiceaccount)
        If MakeUpper = True Then invoiceaccount = UCase$(invoiceaccount)
        If Len(LimitGroupTo) > 0 Then
            If MvSearch(LimitGroupTo, invoiceaccount, VM) = 0 Then GoTo NextRecord
        End If
        Set tbS = SeekRs(StrGroupFld, invoiceaccount, UsrName & "MonthlySchedule")
        Select Case tbS.EOF
        Case True
            tbS.AddNew
            tbS.Fields(StrGroupFld) = invoiceaccount
            tbS.Fields(fName) = invoicetotal
            tbS.Fields("totals") = invoicetotal
            UpdateRs tbS
        Case Else
            oldValue = MyRN(tbS.Fields(fName))
            oldValue = Val(oldValue) + Val(invoicetotal)
            oldValue = ProperAmount(oldValue)
            tbS.Fields(fName) = oldValue
            tbS.Fields("totals") = Val(MyRN(tbS.Fields("totals"))) + Val(invoicetotal)
            tbS.Fields("totals") = ProperAmount(MyRN(tbS.Fields("totals")))
            UpdateRs tbS
        End Select
        'Set tbS = SeekRs(StrGroupFld, "Totals", UsrName & "MonthlySchedule")
        'Select Case tbS.EOF
        'Case True
        '    tbS.AddNew
        '    tbS.Fields(StrGroupFld) = "Totals"
        '    tbS.Fields(fName) = ProperAmount(invoicetotal)
        '    tbS.Fields("totals") = invoicetotal
        '    tbS.Fields("totals") = ProperAmount(myrn(tbS.Fields("totals")))
        '    UpdateRs tbS
        'Case Else
        '    tbS.Fields(fName) = Val(myrn(tbS.Fields(fName))) + Val(invoicetotal)
        '    tbS.Fields(fName) = ProperAmount(myrn(tbS.Fields(fName)))
        '    tbS.Fields("totals") = Val(myrn(tbS.Fields("totals"))) + Val(invoicetotal)
        '    tbS.Fields("totals") = ProperAmount(myrn(tbS.Fields("totals")))
        '    UpdateRs tbS
        'End Select
NextRecord:
        DoEvents
        tbF.MoveNext
    Next
    tbF.Close
    ProgBarClose frmObject.progBar
    If RemoveEmptyColumns = True Then
        StatusMessage frmObject, "Removing columns that are empty, please be patient..."
        RemoveEmptyColumns UsrName & "MonthlySchedule"
    End If
    sMonths = FieldNames(UsrName & "MonthlySchedule")
    If Len(LastQuery) = 0 Then
        If ShowTotalOnly = True Then
            ViewSQLNew "select `" & StrGroupFld & "`,Totals from `" & UsrName & "MonthlySchedule` order by `" & StrGroupFld & "`", LstView, StrGroupFld & ",Totals"
        Else
            ViewSQLNew UsrName & "MonthlySchedule", LstView, sMonths
        End If
    Else
        ViewSQLNew LastQuery, LstView, sMonths
    End If
    LstView.ColumnHeaders(1).Text = ProperCase(NewColumnName)
    sMonths = LstViewColNames(LstView)
    fTot = MvCount(sMonths, ",")
    For fCnt = 2 To fTot
        LstViewSumColumns LstView, True, MvField(sMonths, fCnt, ",")
    Next
    Call LstViewAutoResize(LstView)
    LstView.Refresh
    StatusMessage frmObject, LstView.ListItems.Count & " record(s) selected.", 4
    DeleteTables UsrName & "MonthlySchedule"
    Err.Clear
End Sub
Function DotValue(ByVal StrValue As String) As String
    On Error Resume Next
    Dim s_size As Long
    Dim s_cents As String
    Dim s_firstpart As Long
    Dim s_numbers As String
    StrValue = Trim$(StrValue)
    If Len(StrValue) = 0 Then
        DotValue = "0.00"
    Else
        Select Case StrValue
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            StrValue = "00" & StrValue
        End Select
        If InStr(1, StrValue, ".") = 0 Then
            s_size = Len(StrValue)
            s_cents = Right$(StrValue, 2)
            s_firstpart = s_size - 2
            s_numbers = Left$(StrValue, s_firstpart)
            DotValue = s_numbers & "." & s_cents
        Else
            DotValue = StrValue
        End If
    End If
    Err.Clear
End Function
Public Sub ImportEntityBankDetails(frmObj As Form, ByVal strFile As String)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim rsStr As String
    Dim fLines() As String
    Dim fContents As String
    Dim tb As New ADODB.Recordset
    Dim nameDetails As String
    Dim surname As String
    Dim firstname As String
    Dim Title As String
    Dim eType As String
    Dim accountDetails As String
    Dim accountStatus As String
    Dim AccountType As String
    Dim AccountNumber As String
    Dim BankName As String
    Dim BranchName As String
    Dim branchCode As String
    Dim effectiveDate As String
    Dim nLast As String
    Dim accT As Long
    Dim accC As Long
    Dim accE As Long
    Dim sEntityType As String
    sEntityType = UCase$(FileToken(strFile, "fo"))
    Execute "delete from `Entity Bank Details` where EntityType = '" & EscIn(sEntityType) & "';"
    Set tb = OpenRs("Entity Bank Details", , 1)
    fContents = FileData(strFile)
    rsTot = StrParse(fLines, fContents, vbNewLine)
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing " & FileToken(strFile, "fo")
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        rsStr = RTrim$(fLines(rsCnt))
        Select Case rsStr
        Case "-----------------------------------------------------------ENTITY DETAILS-------"
            ' this is a start of a new record
            nameDetails = fLines(rsCnt + 3)
            surname = Trim$(Left$(nameDetails, 32))
            firstname = Trim$(Mid$(nameDetails, 34, 32))
            Title = Trim$(Mid$(nameDetails, 67, 6))
            eType = MVLastItem(nameDetails, " ")
            accC = rsCnt + 8
            accE = rsCnt + 5000
            For accT = accC To accE
                accountDetails = fLines(accT)
                accountStatus = UCase$(Trim$(Mid$(accountDetails, 114, 6)))
                Select Case accountStatus
                Case "ACTIVE", "PENDNG", "INACT", "REJECT", "TOBEAC"
                    AccountType = Trim$(Mid$(accountDetails, 1, 25))
                    AccountNumber = Trim$(Mid$(accountDetails, 27, 15))
                    BankName = Trim$(Mid$(accountDetails, 43, 32))
                    BranchName = Trim$(Mid$(accountDetails, 75, 32))
                    branchCode = Trim$(Mid$(accountDetails, 107, 6))
                    effectiveDate = Trim$(Mid$(accountDetails, 121))
                    BranchName = Trim$(Replace$(BranchName, "*", ""))
                    nLast = MVLastItem(BranchName, " ")
                    If IsNumeric(nLast) = True Then
                        BranchName = Trim$(Replace$(BranchName, nLast, ""))
                    End If
                    tb.AddNew
                    tb.Fields("surname") = surname
                    tb.Fields("FirstNames") = firstname
                    tb.Fields("title") = Title
                    tb.Fields("AccountType") = MvField(AccountType, 1, " ")
                    tb.Fields("AccountNumber") = AccountNumber
                    tb.Fields("BankName") = BankName
                    tb.Fields("BranchName") = BranchName
                    tb.Fields("branchCode") = branchCode
                    tb.Fields("Status") = accountStatus
                    tb.Fields("effectiveDate") = effectiveDate
                    tb.Fields("FullName") = Trim$(firstname & " " & surname)
                    tb.Fields("EntityType") = sEntityType
                    UpdateRs tb
                Case Else
                    Exit For
                End Select
            Next
        End Select
        DoEvents
    Next
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Set tb = Nothing
    Err.Clear
End Sub
Public Function DaoCountRecords(ByVal Dbase As String, ByVal Table As String) As Long
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Set db = DAO.OpenDatabase(Dbase)
    Set tb = db.OpenRecordset(Table)
    DaoCountRecords = tb.RecordCount
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Function
Public Function EstablishConnection(WinClient As MSWinsockLib.Winsock, ByVal StrComputer As String) As Boolean
    On Error Resume Next
    If Not WinClient.State = sckConnected Then WinClient.Close
    If Not WinClient.State = sckConnected Then
        Call WinClient.Connect(StrComputer, 9456)
        Do While WinClient.State <> sckConnected
            DoEvents
            If WinClient.State = sckError Then Exit Do
        Loop
        If WinClient.State = sckError Then
            WinClient.Close
            EstablishConnection = False
        Else
            If WinClient.State = sckConnected Then
                EstablishConnection = True
            Else
                EstablishConnection = False
            End If
        End If
    Else
        EstablishConnection = True
    End If
    Err.Clear
End Function
Public Sub WinsockSendData(WinClient As MSWinsockLib.Winsock, ByVal sAction As String, Optional ByVal sFields As String = "", Optional ByVal sValues As String = "")
    On Error Resume Next
    Dim strData As String
    strData = StringToMv(KM, sAction, sFields, sValues)
    If WinClient.State = sckConnected Then
        WinClient.SendData strData
    End If
    Err.Clear
End Sub
Function StringsConcat(ParamArray items()) As String
    On Error Resume Next
    Dim Item As Variant
    Dim NewString As String
    NewString = ""
    For Each Item In items
        NewString = Concat(NewString, CStr(Item))
    Next
    StringsConcat = NewString
    Set Item = Nothing
    Err.Clear
End Function
Function Concat(ByVal dest As String, ByVal Source As String) As String
    On Error Resume Next
    Dim sl As Long
    Dim dL As Long
    Dim NL As Long
    Dim sN As String
    Const cI As Long = 50000
    sN = dest
    sl = Len(Source)
    dL = Len(dest)
    NL = dL + sl
    Select Case NL
    Case Is >= dL
        Select Case sl
        Case Is > cI
            sN = sN & Space$(sl)
        Case Else
            sN = sN & Space$(sl + 1)
        End Select
    End Select
    Mid$(sN, dL + 1, sl) = Source
    Concat = Left$(sN, NL)
    Err.Clear
End Function
Function RemoveDotNotation(ByVal fldName As String) As String
    On Error Resume Next
    Dim dotPos As Integer
    dotPos = InStr(fldName, ".")
    Select Case dotPos
    Case 0
        RemoveDotNotation = fldName
    Case Else
        RemoveDotNotation = Mid$(fldName, dotPos + 1)
    End Select
    Err.Clear
End Function
Sub ArrayTrimItems(varArray() As String)
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
Public Function StringAdd(ByVal Strdest As String, ByVal Straddstring As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim NewString As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    NewString = Concat(Strdest, Straddstring)
    NewString = Concat(NewString, Delim)
    StringAdd = NewString
    Err.Clear
End Function
Public Function LstViewColNames(LstView As ListView) As String
    On Error Resume Next
    Dim strHead As String
    Dim strName As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    strHead = ""
    clsColTot = LstView.ColumnHeaders.Count
    For clsColCnt = 1 To clsColTot
        strName = LstView.ColumnHeaders(clsColCnt).Text
        Select Case clsColCnt
        Case clsColTot
            strHead = Concat(strHead, strName)
        Case Else
            strHead = StringAdd(strHead, strName, ",")
        End Select
    Next
    LstViewColNames = strHead
    Err.Clear
End Function
Public Function MvCount(ByVal StringMv As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim xNew() As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    xNew = Split(StringMv, Delim)
    MvCount = UBound(xNew) + 1
    Err.Clear
End Function
Public Sub LstViewSaveReport(LstView As ListView, ByVal ReportName As String, Optional ByVal ReportTable As String = "MyReports")
    On Error Resume Next
    Dim colNames As String
    Dim colAlignment As String
    Dim spLine() As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim repData As String
    Dim curAlign As String
    Dim repLine As String
    If LstView.ListItems.Count = 0 Then Exit Sub
    colNames = LstViewColNames(LstView)
    ' how many columns do we have
    rsTot = MvCount(colNames, ",")
    ' read the alignment of the columns
    colAlignment = ""
    For rsCnt = 1 To rsTot
        Select Case LstView.ColumnHeaders(rsCnt).Alignment
        Case 2 ' centre
            curAlign = "C"
        Case 0 ' left
            curAlign = "L"
        Case 1 ' left
            curAlign = "R"
        End Select
        If rsCnt = rsTot Then
            colAlignment = colAlignment & curAlign
        Else
            colAlignment = colAlignment & curAlign & ","
        End If
    Next
    If TableExists(ReportTable) = False Then
        Execute "CREATE TABLE " & ReportTable & " (ID VARCHAR(255),User VARCHAR(255),ColumnNames LONGTEXT,ColumnAlignment LONGTEXT,Contents LONGTEXT, Key (ID)) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;"
        DoEvents
    End If
    repData = ""
    rsTot = LstView.ListItems.Count
    For rsCnt = 1 To rsTot
        spLine = LstViewGetRow(LstView, rsCnt)
        repLine = MvFromArray(spLine, FM)
        If rsCnt = rsTot Then
            repData = repData & repLine
        Else
            repData = repData & repLine & RM
        End If
    Next
    Dim Record(5) As String
    Record(1) = ReportName
    Record(2) = colNames
    Record(3) = colAlignment
    Record(4) = repData
    Record(5) = UserName
    myValues = MvFromArray(Record, FM)
    myFields = "ID,ColumnNames,ColumnAlignment,Contents,UserName"
    WriteRecordMv ReportTable, "id", ReportName, myFields, myValues, FM
    Err.Clear
End Sub
Public Function LstViewGetRow(LstView As ListView, ByVal idx As Long) As Variant
    On Error Resume Next
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    clsColTot = LstView.ColumnHeaders.Count
    ReDim retarray(clsColTot)
    retarray(1) = LstView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = LstView.ListItems(idx).SubItems(clsColCnt)
    Next
    LstViewGetRow = retarray
    Err.Clear
End Function
Public Function MvRemoveItems(ByVal MvString As String, ByVal Delim As String, ParamArray items()) As String
    On Error Resume Next
    Dim Item As Variant
    Dim spItems() As String
    Dim StrNew As String
    Dim spItemTot As Long
    Dim spItemCnt As Long
    Call StrParse(spItems, MvString, Delim)
    spItemTot = UBound(spItems)
    For Each Item In items
        For spItemCnt = 1 To spItemTot
            If LCase$(spItems(spItemCnt)) = LCase$(CStr(Item)) Then
                spItems(spItemCnt) = "{}"
            End If
        Next
    Next
    StrNew = ""
    For spItemCnt = 1 To spItemTot
        If spItems(spItemCnt) <> "{}" Then
            StrNew = StringsConcat(StrNew, spItems(spItemCnt), Delim)
        End If
    Next
    MvRemoveItems = RemoveDelim(StrNew, Delim)
    Err.Clear
End Function
Public Sub TreeViewLoadTbFields(ByVal TbName As String, ByVal TbFlds As String, TreeV As TreeView, Optional ByVal Image As String = "closed", Optional ByVal SelImage As String = "closed", Optional ByVal TagFld As String = "", Optional ByVal SortOrder As String = "", Optional ByVal Delim As String = "\", Optional ByVal SortExpandTree As Boolean = False, Optional ByVal WhereCriteria As String = "")
    On Error Resume Next
    Dim mRs As ADODB.Recordset
    Dim varColN() As String
    Dim st1 As String
    Dim st2 As String
    Dim sql As String
    Dim varColTot As Long
    Dim varColCnt As Long
    Dim rsTot As Long
    Dim rsCnt As Long
    sql = "select " & SQLQuote(TbFlds) & " from `" & TbName & "`"
    If Len(WhereCriteria) > 0 Then sql = sql & " " & WhereCriteria
    If Len(SortOrder) > 0 Then sql = sql & " order by " & SortOrder
    Set mRs = OpenRs(sql)
    rsTot = AffectedRecords
    For rsCnt = 1 To rsTot
        st1 = ""
        st2 = ""
        Call StrParse(varColN, TbFlds, ",")
        varColTot = UBound(varColN)
        For varColCnt = 1 To varColTot
            varColN(varColCnt) = MyRN(mRs.Fields(varColN(varColCnt)))
            varColN(varColCnt) = ProperCase(varColN(varColCnt))
            Select Case varColCnt
            Case varColTot
                st1 = StringsConcat(st1, varColN(varColCnt))
            Case Else
                st1 = StringsConcat(st1, varColN(varColCnt), Delim)
            End Select
        Next
        If Len(TagFld) > 0 Then
            st2 = MyRN(mRs.Fields(TagFld))
        End If
        TreeViewAddPathWithKey TreeV, st1, Image, SelImage, st2, Delim
        mRs.MoveNext
    Next
    mRs.Close
    Set mRs = Nothing
    If SortExpandTree = True Then
        varColTot = TreeV.Nodes.Count
        For varColCnt = 1 To varColTot
            TreeV.Nodes(varColCnt).Sorted = True
            TreeV.Nodes(varColCnt).Expanded = True
        Next
    End If
    Err.Clear
End Sub
Public Function DaoTableFieldAutoIncrement(ByVal Dbase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Dim att As Integer
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Fields.Count - 1
    fL = ""
    For fC = 0 To fT
        att = db.TableDefs(TbName).Fields(fC).Attributes
        If att >= 16 And att <= 100 Then
            fL = fL & db.TableDefs(TbName).Fields(fC).Name & ","
        End If
    Next
    DaoTableFieldAutoIncrement = RemoveDelim(fL, ",")
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Sub LstViewSumRest(lstReport As ListView)
    On Error Resume Next
    Dim xCols() As String
    Dim sCols As String
    Dim rsCnt As Long
    Dim rsTot As Long
    sCols = LstViewColNames(lstReport)
    Call StrParse(xCols, sCols, ",")
    rsTot = UBound(xCols)
    For rsCnt = 2 To rsTot
        LstViewSumColumns lstReport, True, xCols(rsCnt)
    Next
    Err.Clear
End Sub
Public Function SQLDateToNormal(ByVal strDate As String) As String
    On Error Resume Next
    Dim sYYYY As String
    Dim smm As String
    Dim sdd As String
    SQLDateToNormal = strDate
    If MvCount(strDate, "-") = 3 Then
        sYYYY = StringPart(strDate, 1, "-")
        smm = StringPart(strDate, 2, "-")
        sdd = StringPart(strDate, 3, "-")
        SQLDateToNormal = sdd & "/" & smm & "/" & sYYYY
    End If
    Err.Clear
End Function
Public Function DaoTableIndexes(ByVal Dbase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Indexes.Count - 1
    fL = ""
    For fC = 0 To fT
        fN = db.TableDefs(TbName).Indexes(fC).Name
        If fN <> "PrimaryKey" Then
            Select Case fC
            Case fT
                fL = fL & fN
            Case Else
                fL = fL & fN & ","
            End Select
        End If
    Next
    DaoTableIndexes = fL
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Public Function DaoTablePrimaryIndexes(ByVal Dbase As String, ByVal TbName As String) As String '
    On Error Resume Next
    Dim db As DAO.Database
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Indexes.Count - 1
    For fC = 0 To fT
        If db.TableDefs(TbName).Indexes(fC).Primary = True Then
            fN = db.TableDefs(TbName).Indexes(fC).Fields(0).Name
            Exit For
        End If
    Next
    DaoTablePrimaryIndexes = fN
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Public Function DaoTableFieldNames(ByVal Dbase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Fields.Count - 1
    fL = ""
    For fC = 0 To fT
        fN = db.TableDefs(TbName).Fields(fC).Name
        Select Case fC
        Case fT
            fL = fL & fN
        Case Else
            fL = fL & fN & ","
        End Select
    Next
    DaoTableFieldNames = fL
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Public Function DaoTableFieldTypes(ByVal Dbase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Dim fType As Long
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Fields.Count - 1
    fL = ""
    For fC = 0 To fT
        fType = db.TableDefs(TbName).Fields(fC).Type
        Select Case fType
        Case dbBigInt
            fN = "bigint"
        Case dbLongBinary
            fN = "longbinary"
        Case dbChar
            fN = "char"
        Case dbDecimal
            fN = "decimal"
        Case dbFloat
            fN = "float"
        Case dbGUID
            fN = "guid"
        Case dbTime
            fN = "time"
        Case dbTimeStamp
            fN = "timestamp"
        Case dbNumeric
            fN = "numeric"
        Case dbVarBinary
            fN = "varbinary"
        Case dbBoolean
            fN = "boolean"
        Case dbByte
            fN = "byte"
        Case dbInteger
            fN = "integer"
        Case dbLong
            fN = "long"
        Case dbCurrency
            fN = "currency"
        Case dbSingle
            fN = "single"
        Case dbDouble
            fN = "double"
        Case dbDate
            fN = "date"
        Case dbText
            fN = "text"
        Case dbMemo
            fN = "memo"
        End Select
        Select Case fC
        Case fT
            fL = fL & fN
        Case Else
            fL = fL & fN & ","
        End Select
    Next
    DaoTableFieldTypes = fL
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Public Function DaoTableFieldSizes(ByVal Dbase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = DAO.OpenDatabase(Dbase)
    fT = db.TableDefs(TbName).Fields.Count - 1
    fL = ""
    For fC = 0 To fT
        fN = db.TableDefs(TbName).Fields(fC).Size
        If fN = "0" Then fN = ""
        Select Case fC
        Case fT
            fL = fL & fN
        Case Else
            fL = fL & fN & ","
        End Select
    Next
    DaoTableFieldSizes = fL
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Sub PrintExcel(ByVal SavePath As String, ByVal StrCaption As String, lstReport As ListView, Optional FootNote As String = "", Optional Italic As Boolean = True, Optional FixPipe As Boolean = True, Optional ToPrinter As Boolean = False, Optional FitPage As Boolean = False, Optional intFontSize As Integer = 8)
    On Error Resume Next
    Dim xFile As String
    Dim xPath As String
    If lstReport.ListItems.Count = 0 Then Exit Sub
    xFile = SavePath & "\" & StrCaption & ".xls"
    xFile = FileName_Validate(xFile)
    xPath = FileToken(xFile, "p")
    FootNote = ProperCase(Province & " Department Of " & Department)
    If DirExists(xPath) = False Then MakeDirectory xPath
    LstViewToWorksheet lstReport, xFile, StrCaption, FootNote, "", "", True, True, , , FitPage, , intFontSize, Italic
    DoEvents
    If FixPipe = True Then ExcelFindReplace xFile, "|", Chr$(10)
    DoEvents
    If FileExists(xFile) = True Then
        If ToPrinter = True Then
            Call ViewFile(xFile, "print")
        Else
            Call ViewFile(xFile, "open")
        End If
    End If
    Err.Clear
End Sub
Public Function DirExists(ByVal Sdirname As String) As Boolean
    On Error Resume Next
    Dim sDir As String
    DirExists = False
    sDir = Dir$(Sdirname, vbDirectory)
    If Len(sDir) > 0 Then DirExists = True
    Err.Clear
End Function
Public Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(sporiginal, strOriginalMv, Delim)
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = Concat(sLine, sporiginal(spCnt))
        Else
            sLine = StringsConcat(sLine, sporiginal(spCnt), Delim)
        End If
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Public Function ComputerName() As String
    On Error Resume Next
    Dim compname As String
    compname = Space$(255)
    Call GetComputerName(compname, 255)  ' get the computer's name
    ComputerName = Left$(compname, InStr(compname, vbNullChar) - 1)
    ComputerName = MvField(ComputerName, 1, ".")
    Err.Clear
End Function
Public Sub MakeDirectory(ByVal Sdirectory As String)
    On Error GoTo CreateDirectory_ErrorHandler
    CreateNestedDirectory Sdirectory
    Err.Clear
    Exit Sub
CreateDirectory_ErrorHandler:
    Select Case Err
    Case 0
    Case 75
    Case Else
        MsgBox "Directory Name : " & Sdirectory & vbCr & vbCr & "Error " & VBA.CStr(Err) & ":" & "  " & Error & vbCr & "Please check your drive and disk." & vbCr & vbCr & "Directory Name" & Sdirectory, vbOKOnly + vbExclamation, "Create Directory"
    End Select
    Err.Clear
End Sub
Public Sub LstViewToWorksheet(LstView As ListView, ByVal strFile As String, Optional ByVal StrHeader As String = "", Optional ByVal LeftFooter As String = "", Optional ByVal CenterFooter As String = "", Optional ByVal strTab As String = "", Optional ByVal boolNew As Boolean = True, Optional ByVal boolShow As Boolean = False, Optional ByVal pOrientation As String = "Landscape", Optional boolCenterHeading As Boolean = False, Optional boolFitToPage As Boolean = False, Optional boolIncreaseTopMargin As Boolean = False, Optional intFontSize As Integer = 8, Optional bFontItalic As Boolean = True, Optional ZoomTo As Integer = 100, Optional ShowProcess As Boolean = True)
    On Error Resume Next
    Err.Clear
    Exit Sub
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim ColumnNames As String
    Dim spColumns() As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xPageSetUp As Excel.PageSetup
    Dim bExist As Boolean
    Dim lngNext As Integer
    Dim bFound As Boolean
    Dim sheetName As String
    Dim DbExcel As DAO.Database
    Dim RsExcel As DAO.Recordset
    Dim spLine() As String
    Dim colCnt As Long
    Dim colTot As Long
    Dim xPath As String
    strFile = FileName_Validate(strFile)
    xPath = FileToken(strFile, "p")
    If DirExists(xPath) = False Then CreateNestedDirectory xPath
    bExist = FileExists(strFile)
    If boolNew = True Then
        If bExist = True Then
            Kill strFile
            bExist = False
        End If
    End If
    If DirExists(FileToken(strFile, "p")) = False Then
        MakeDirectory FileToken(strFile, "p")
    End If
    Set xlApp = New Excel.Application
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.WindowState = Excel.xlMinimized
    xlApp.Visible = False
    If bExist = True Then
        xlApp.Workbooks.Open strFile
    Else
        xlApp.Workbooks.Add
    End If
    Set xlBook = xlApp.ActiveWorkbook
    lngNext = 1
    Do Until bFound = True
        Set xlSheet = xlApp.Worksheets("Sheet" & lngNext)
        If TypeName(xlSheet) = "Nothing" Then
            bFound = False
            lngNext = lngNext + 1
        Else
            bFound = True
        End If
    Loop
    If Len(strTab) > 0 Then
        xlSheet.Name = ProperCase(ExcelCorrectSheetName(strTab))
    End If
    ColumnNames = LstViewColNames(LstView)
    colTot = StrParse(spColumns, ColumnNames, ",")
    With xlSheet
        .Activate
        .Cells.NumberFormat = "General"
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = intFontSize
        For colCnt = 1 To colTot
            .Cells(1, colCnt).Value = spColumns(colCnt)
        Next
        .SaveAs strFile
        sheetName = .Name
    End With
    ' Close the Workbook
    xlBook.Close
    ' Close Microsoft Excel with the Quit method.
    xlApp.Quit
    ' open the workbook and treat it as an access database
    Set DbExcel = DAO.OpenDatabase(strFile, False, False, "Excel 8.0; HDR=YES;")
    Set RsExcel = DbExcel.OpenRecordset(sheetName & "$")
    rsTot = LstView.ListItems.Count
    For rsCnt = 1 To rsTot
        spLine = LstViewGetRow(LstView, rsCnt)
        RsExcel.AddNew
        For colCnt = 1 To colTot
            RsExcel(colCnt - 1) = spLine(colCnt)
        Next
        RsExcel.Update
    Next
    RsExcel.Close
    DbExcel.Close
    ' reopen excel and format report
    Set xlApp = New Excel.Application
    xlApp.DisplayAlerts = False
    xlApp.WindowState = Excel.xlMinimized
    xlApp.ScreenUpdating = False
    xlApp.Visible = False
    xlApp.Workbooks.Open strFile
    Set xlBook = xlApp.ActiveWorkbook
    Set xlSheet = xlApp.Worksheets(sheetName)
    With xlSheet
        .Activate
        .Cells.NumberFormat = "General"
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = intFontSize
        For colCnt = 1 To colTot
            .Cells(1, colCnt).Value = spColumns(colCnt)
            .Cells(1, colCnt).Interior.ColorIndex = 15
            .Cells(1, colCnt).Font.Bold = True
            .Cells(1, colCnt).Borders.Weight = Excel.XlBorderWeight.xlThin
            .Cells(1, colCnt).Interior.Pattern = Excel.xlSolid
            .Columns(colCnt).AutoFit
            If LstView.ColumnHeaders(colCnt).Alignment = lvwColumnCenter Then
                .Columns(colCnt).HorizontalAlignment = Excel.xlCenter
            ElseIf LstView.ColumnHeaders(colCnt).Alignment = lvwColumnRight Then
                .Columns(colCnt).HorizontalAlignment = Excel.xlRight
            End If
        Next
        rsTot = LstView.ListItems.Count + 1
        .Cells(1, 1).Resize(rsTot, colTot).Borders.Weight = Excel.XlBorderWeight.xlThin
        .Cells(1, 1).Resize(rsTot, colTot).Font.Italic = bFontItalic
        .Cells(1, 1).Resize(rsTot, colTot).VerticalAlignment = Excel.xlTop
        If boolCenterHeading = True Then
            .Rows(1).HorizontalAlignment = xlCenter
        End If
        .Rows.AutoFit
    End With
    GoSub DoPageSetup
    xlBook.Save
    ' Close the Workbook
    xlBook.Close
    ' Close Microsoft Excel with the Quit method.
    xlApp.Quit
    ' Release the objects.
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Err.Clear
    Exit Sub
DoPageSetup:
    Set xPageSetUp = xlSheet.PageSetup
    With xPageSetUp
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
        .PrintArea = ""
        .CenterHeader = ""
        .LeftHeader = "&B&I&" & Quote & "Tahoma" & Quote & "&10" & Replace$(StrHeader, "&", "&&")
        .RightHeader = ""
        .LeftFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & Replace$(LeftFooter, "&", "&&")
        .CenterFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & Replace$(CenterFooter, "&", "&&")
        .RightFooter = "&" & Quote & "Tahoma" & Quote & "&8" & "Page &P of &N"
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintNotes = False
        .CenterHorizontally = False
        .CenterVertically = False
        Select Case LCase$(pOrientation)
        Case "landscape"
            .Orientation = Excel.xlLandscape
        Case Else
            .Orientation = Excel.xlPortrait
        End Select
        .Draft = False
        .PaperSize = Excel.xlPaperA4
        .FirstPageNumber = Excel.xlAutomatic
        .Order = Excel.xlDownThenOver
        .BlackAndWhite = False
        If boolFitToPage = True Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        Else
            .Zoom = ZoomTo
        End If
        If boolIncreaseTopMargin = True Then
            .TopMargin = xlApp.InchesToPoints(1.37795275590551)
        End If
        DoEvents
    End With
    Err.Clear
    Return
    Err.Clear
End Sub
Sub ExcelFindReplace(ByVal strFile As String, ByVal StrFind As String, ByVal StrReplaceWith As String)
    On Error Resume Next
    Err.Clear
    Exit Sub
    Dim exlApp As Excel.Application
    Dim XLWkb As Excel.Workbook
    Dim xlWks As Excel.Worksheet
    Dim sheetNames As String
    Dim spSheetNames() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim spStr As String
    Set exlApp = New Excel.Application
    If TypeName(exlApp) = "Nothing" Then
    Err.Clear
        Exit Sub
    End If
    sheetNames = ExcelSheetNames(strFile, VM)
    spTot = StrParse(spSheetNames, sheetNames, VM)
    exlApp.DisplayAlerts = False
    exlApp.ScreenUpdating = False
    exlApp.Workbooks.Open strFile
    exlApp.WindowState = xlMinimized
    exlApp.Visible = False
    Set XLWkb = exlApp.ActiveWorkbook
    XLWkb.Activate
    For spCnt = 1 To spTot
        spStr = spSheetNames(spCnt)
        Set xlWks = XLWkb.Worksheets(spStr)
        xlWks.Cells.Cells.Replace What:=StrFind, Replacement:=StrReplaceWith, LookAt:=Excel.xlPart, SearchOrder:=Excel.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        xlWks.Columns.AutoFit
        xlWks.Rows.AutoFit
    Next
    XLWkb.Save
    exlApp.Quit
    DoEvents
    Set exlApp = Nothing
    Set XLWkb = Nothing
    Set xlWks = Nothing
    Err.Clear
End Sub
Public Function FileExists(ByVal Filename As String) As Boolean
    On Error Resume Next
    FileExists = False
    If Len(Filename) = 0 Then
    Err.Clear
        Exit Function
    End If
    FileExists = IIf(Dir$(Filename) <> "", True, False)
    Err.Clear
End Function
Public Function ViewFile(ByVal Filename As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Boolean
    On Error Resume Next
    Dim R As Long
    R = lngStartDoc(Filename, Operation, WindowState)
    If R <= 32 Then
        ' there was an error
        Beep
        MsgBox "An error occurred while opening your document." & vbCr & "The possibility is that the selected entry does not have" & vbCr & "a link in the registry to open it with.", vbOKOnly + vbExclamation, "Viewer Error"
        ViewFile = False
    Else
        ViewFile = True
        Pause 1
    End If
    Err.Clear
End Function
Sub CreateNestedDirectory(ByVal StrCompletePath As String)
    On Error Resume Next
    Dim spPaths() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim curPath As String
    Call StrParse(spPaths, StrCompletePath, "\")
    spTot = UBound(spPaths)
    For spCnt = 1 To spTot
        curPath = MvFromMv(StrCompletePath, 1, spCnt, "\")
        If DirExists(curPath) = False Then
            MkDir curPath
        End If
    Next
    Err.Clear
End Sub
Function ExcelCorrectSheetName(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = Replace$(StrValue, ":", "")
    StrValue = Replace$(StrValue, ",", "")
    StrValue = Replace$(StrValue, "/", "")
    StrValue = Replace$(StrValue, "\", "")
    StrValue = Replace$(StrValue, "?", "")
    StrValue = Replace$(StrValue, "*", "")
    StrValue = Replace$(StrValue, "[", "")
    StrValue = Replace$(StrValue, "]", "")
    ExcelCorrectSheetName = Left$(Trim$(StrValue), 31)
    Err.Clear
End Function
Public Function ExcelSheetNames(ByVal strFile As String, Optional Delimiter As String = ";") As String
    On Error Resume Next
    Err.Clear
    Exit Function
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim xlApp As Excel.Application
    Dim XLWkb As Excel.Workbook
    Dim sRslt As String
    Set xlApp = New Excel.Application
    If TypeName(xlApp) = "Nothing" Then
    Err.Clear
        Exit Function
    End If
    xlApp.Visible = False
    xlApp.Workbooks.Open strFile
    Set XLWkb = xlApp.ActiveWorkbook
    XLWkb.Activate
    Set XLWkb = xlApp.ActiveWorkbook
    rsTot = XLWkb.Worksheets.Count
    sRslt = ""
    For rsCnt = 1 To rsTot
        sRslt = sRslt & XLWkb.Worksheets(rsCnt).Name & Delimiter
    Next
    sRslt = RemoveDelim(sRslt, Delimiter)
    XLWkb.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set XLWkb = Nothing
    ExcelSheetNames = sRslt
    Err.Clear
End Function
Function lngStartDoc(ByVal Docname As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Long
    On Error Resume Next
    Dim Scr_hDC As Long
    Dim sDir As String
    sDir = FileToken(Docname, "d")
    Scr_hDC = GetDesktopWindow()
    lngStartDoc = ShellExecute(Scr_hDC, Operation, Docname, "", sDir, WindowState)
    Err.Clear
End Function
Public Sub Pause(ByVal nSecond As Double)
    On Error Resume Next
    ' call pause(2)      delay for 2 seconds
    Dim t0 As Double
    t0 = Timer
    Do While Timer - t0 < nSecond
        DoEvents
        ' if we cross midnight, back up one day
        If Timer < t0 Then
            t0 = t0 - CLng(24) * CLng(60) * CLng(60)
        End If
    Loop
    Err.Clear
End Sub
Public Function SetDate() As Boolean
    On Error Resume Next
    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()
    If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "dd/MM/yyyy") = False Then
        Call MyPrompt("The system date could not be changed to dd/MM/yyyy format, please change it in the control panel.", "o", "e", "System Date")
        SetDate = False
    Err.Clear
        Exit Function
    Else
        SetDate = True
        PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    End If
    Err.Clear
End Function
Function MyPrompt(ByVal StrMsg As String, Optional ByVal strButton As String = "o", Optional ByVal StrIcon As String = "e", Optional ByVal StrHeading As String = "") As Variant
    On Error Resume Next
    ' button can be any of
    ' ync - yesnocancel, c - cancel, o - ok, oc - okcancel, rc - retrycancel and yn - yesno
    ' and ari - abortretryignore, bc - backclose, bnc - backnextclose
    ' bns - backnextsnooze, nc - nextclose, sc - searchclose, toc - tipsoptionsclose, yanc - yesallnocancel
    ' icon can be any of
    ' i - information, w - warning, c - critical, t - tip, q - query
    ' mode can be any of
    ' ad - autodown, ma - modal, me - modeless
    Dim isCheck As Long
    Dim Mode As Long
    Dim Button As Long
    Dim Icon As Long
    ' see if excel is already running
    If Len(StrHeading) = 0 Then
        StrHeading = App.Title
    End If
    isCheck = 0
    Mode = vbApplicationModal
    Select Case LCase$(strButton)
    Case "ync"
        Button = vbYesNoCancel
    Case "c"
        Button = vbCancel
    Case "o"
        Button = vbOKOnly
    Case "oc"
        Button = vbOKCancel
    Case "rc"
        Button = vbRetryCancel
    Case "yn"
        Button = vbYesNo
    Case "ari"
        Button = vbAbortRetryIgnore
    End Select
    Select Case LCase$(StrIcon)
    Case "i", "t"
        Icon = vbInformation
    Case "w", "e"
        Icon = vbExclamation
    Case "c"
        Icon = vbCritical
    Case "q"
        Icon = vbQuestion
    End Select
    MyPrompt = MsgBox(StrMsg, Button + Icon + Mode, StrHeading)
    Err.Clear
End Function
Function Abbreviate(ByVal StrOriginal As String) As String
    On Error Resume Next
    Dim spParts() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim strRes As String
    strRes = ""
    spParts = Split(StrOriginal, " ")
    spTot = UBound(spParts)
    For spCnt = 0 To spTot
        strRes = strRes & Left$(spParts(spCnt), 1)
    Next
    Abbreviate = strRes
    Err.Clear
End Function
Public Sub LstViewSwapSort(LstView As ListView, lstHeader As Variant)
    On Error Resume Next
    Select Case LstView.SortOrder
    Case 0
        LstView.SortOrder = 1
    Case Else
        LstView.SortOrder = 0
    End Select
    LstView.SortKey = lstHeader.Index - 1
    ' Set Sorted to True to sort the list.
    LstView.Sorted = True
    LstView.Refresh
    Err.Clear
End Sub
Public Sub LstViewAutoResize(LstView As ListView)
    On Error Resume Next
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If LstView.ListItems.Count = 0 Then
    Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = LstView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        Call SendMessage(LstView.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
    'LstViewResizeMax lstView
    Err.Clear
End Sub
Public Sub LstBoxFromMV(lstObj As Variant, ByVal StringMv As String, Optional ByVal Delim As String = "", Optional ByVal Sclear As String = "", Optional ByVal RemoveDups As String = "")
    On Error Resume Next
    Dim spDel() As String
    Dim spCnt As Long
    Dim wCnt As Long
    Dim xItm As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    If Len(Sclear) = 0 Then
        lstObj.Clear
    End If
    Call StrParse(spDel, StringMv, Delim)
    wCnt = UBound(spDel)
    For spCnt = 1 To wCnt
        xItm = ProperCase(spDel(spCnt))
        If Len(xItm) = 0 Then
            GoTo NextLine
        End If
        Select Case Trim$(UCase$(RemoveDups))
        Case "", "Y"
            LstBoxUpdate lstObj, xItm
        Case Else
            lstObj.AddItem xItm
        End Select
NextLine:
    Next
    Err.Clear
End Sub
Function LstViewColumnPosition(lstReport As ListView, ByVal StrColName As String) As Long
    On Error Resume Next
    Dim xCols As String
    xCols = LstViewColNames(lstReport)
    LstViewColumnPosition = MvSearch(xCols, StrColName, ",")
    Err.Clear
End Function
Public Sub LstViewColIconv(LstView As ListView, ByVal intPos As Integer, ByVal Sconvcode As String)
    On Error Resume Next
    Dim strData As String
    Dim clsRowTot As Long
    Dim clsRowCnt As Long
    Sconvcode = UCase$(Sconvcode)
    clsRowTot = LstView.ListItems.Count
    Select Case intPos
    Case 1
        For clsRowCnt = 1 To clsRowTot
            strData = LstView.ListItems(clsRowCnt).Text
            Select Case Sconvcode
            Case "M"
                strData = ProperAmount(strData)
            Case "D"
                strData = DateIconv(strData)
            Case "C"
                strData = ProperCase(strData)
            End Select
            LstView.ListItems(clsRowCnt).Text = strData
        Next
    Case Else
        For clsRowCnt = 1 To clsRowTot
            strData = LstView.ListItems(clsRowCnt).SubItems(intPos - 1)
            Select Case Sconvcode
            Case "M"
                strData = ProperAmount(strData)
            Case "D"
                strData = DateIconv(strData)
            Case "C"
                strData = ProperCase(strData)
            End Select
            LstView.ListItems(clsRowCnt).SubItems(intPos - 1) = strData
        Next
    End Select
    LstView.Refresh
    Err.Clear
End Sub
Public Sub LstViewToExcelGroupSheets(frmForm As Form, LstView As ListView, ByVal strFile As String, ByVal colName As String, Optional ByVal StrHeader As String = "", Optional ByVal StrLeftFooter As String = "", Optional ByVal StrCenterFooter As String = "", Optional ByVal pOrientation As String = "Landscape", Optional ShowProcess As Boolean = True)
    On Error Resume Next
    LstViewToWorksheetGroupBy frmForm, LstView, strFile, colName, StrHeader, StrLeftFooter, StrCenterFooter, pOrientation, ShowProcess
    Err.Clear
End Sub
Public Function LstViewCheckedToMV(LstView As ListView, ByVal lngPos As Long, Optional ByVal Delim As String = "", Optional bRemoveDuplicates As Boolean = False, Optional bRemoveBlanks As Boolean = False, Optional bRemoveStars As Boolean = True, Optional bRemoveTotals As Boolean = True) As String
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim bOp As Boolean
    Dim lstStr() As String
    Dim retStr As String
    retStr = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    lstTot = LstView.ListItems.Count
    For lstCnt = 1 To lstTot
        bOp = LstView.ListItems(lstCnt).Checked
        Select Case bOp
        Case True
            lstStr = LstViewGetRow(LstView, lstCnt)
            retStr = StringsConcat(retStr, lstStr(lngPos), Delim)
        End Select
    Next
    retStr = RemoveDelim(retStr, Delim)
    If bRemoveTotals = True Then
        retStr = Replace$(retStr, "Totals", "")
    End If
    If bRemoveStars = True Then retStr = Replace$(retStr, "*", "")
    If bRemoveDuplicates = True Then
        retStr = MvRemoveDuplicates(retStr, Delim)
    End If
    If bRemoveBlanks = True Then
        retStr = MvRemoveBlanks(retStr, Delim)
    End If
    LstViewCheckedToMV = retStr
    Err.Clear
End Function
Sub LstViewFilterNew(frmForm As Form, lstReport As ListView, ByVal ColumnName As String, ByVal ColumnValue As String, Optional Remove As Integer = 0)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim xCols As String
    Dim xPos As Long
    Dim spLine() As String
    Dim curValue As String
    xCols = LstViewColNames(lstReport)
    xPos = MvSearch(xCols, ColumnName, ",")
    If xPos = 0 Then Exit Sub
    ColumnValue = LCase$(ColumnValue)
    ColumnValue = MvReplaceItem(ColumnValue, "(blank)", "(none)", VM)
    rsTot = lstReport.ListItems.Count
    ProgBarInit frmForm.progBar, rsTot
    StatusMessage frmForm, "Filtering report..."
    For rsCnt = rsTot To 1 Step -1
        frmForm.progBar.Value = rsCnt
        spLine = LstViewGetRow(lstReport, rsCnt)
        curValue = LCase$(Trim$(spLine(xPos)))
        If curValue = "" Then curValue = "(none)"
        If Remove = 0 Then
            If MvSearch(ColumnValue, curValue, VM) = 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        Else
            If MvSearch(ColumnValue, curValue, VM) > 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        End If
        DoEvents
    Next
    ProgBarClose frmForm.progBar
    LstViewAutoResize lstReport
    StatusMessage frmForm
    Err.Clear
End Sub
Public Sub StatusMessage(Thisform As Form, Optional ByVal Rsmsg As String = "", Optional ByVal pos As Integer = 4)
    On Error Resume Next
    If Val(pos) = 0 Then
        pos = 1
    End If
    Thisform.StatusBar1.Panels.Item(pos) = ProperCase(Rsmsg)
    Thisform.StatusBar1.Refresh
    DoEvents
    Err.Clear
End Sub
Public Function LstViewColMV(LstView As ListView, ByVal intPos As Long, Optional ByVal Delim As String = "", Optional bDistinct As Boolean = False, Optional SumLastOnly As Boolean = False) As String
    On Error Resume Next
    Dim strData As String
    Dim strR As String
    Dim clsRowTot As Long
    Dim clsRowCnt As Long
    Dim colCollection As New Collection
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    strR = ""
    clsRowTot = LstView.ListItems.Count
    Select Case intPos
    Case 1
        For clsRowCnt = 1 To clsRowTot
            strData = LstView.ListItems(clsRowCnt).Text
            If SumLastOnly = True Then strData = MvField(strData, -1, " ")
            If bDistinct = False Then
                strR = StringsConcat(strR, strData, Delim)
            Else
                colCollection.Add strData, strData
            End If
        Next
        strR = RemoveDelim(strR, Delim)
    Case Else
        For clsRowCnt = 1 To clsRowTot
            strData = LstView.ListItems(clsRowCnt).SubItems(intPos - 1)
            If SumLastOnly = True Then strData = MvField(strData, -1, " ")
            If bDistinct = False Then
                strR = StringsConcat(strR, strData, Delim)
            Else
                colCollection.Add strData, strData
            End If
        Next
        strR = RemoveDelim(strR, Delim)
    End Select
    If bDistinct = False Then
        LstViewColMV = strR
    Else
        LstViewColMV = MvFromCollection(colCollection, Delim)
    End If
    Err.Clear
End Function
Function NoCommas(ByVal StrValue As String) As String
    On Error Resume Next
    NoCommas = Replace$(StrValue, ",", "")
    Err.Clear
End Function
Function ProperAmount(ByVal StrValue As String) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsVal As String
    Dim mLeft As String
    Dim mRight As String
    rsStr = ""
    rsTot = Len(StrValue)
    For rsCnt = 1 To rsTot
        rsVal = Mid$(StrValue, rsCnt, 1)
        If InStr(1, "-.0123456789", rsVal) > 0 Then
            rsStr = rsStr & rsVal
        End If
    Next
    rsStr = Trim$(rsStr)
    If Len(rsStr) = 0 Then rsStr = "0.00"
    If InStr(1, rsStr, ".") = 0 Then rsStr = rsStr & ".00"
    StrValue = CDbl(rsStr)
    ProperAmount = Format$(StrValue, "###0.00")
    Err.Clear
End Function
Public Function MakeMoney(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = ProperAmount(StrValue)
    MakeMoney = Format$(StrValue, "#,##0.00")
    Err.Clear
End Function
Public Function DateIconv(ByVal sDate As String) As String
    On Error Resume Next
    DateIconv = sDate
    If Len(sDate) = 0 Then
    Err.Clear
        Exit Function
    End If
    Select Case sDate
    Case Is <> ""
        Select Case IsDate(sDate)
        Case True
            Dim DayZero As Date
            Dim Today As Date
            Dim NumDays As Long
            ' for pick and universe date zero is 31/12/1967
            DayZero = ToDate("31/12/1967")
            Today = CDate(ToDate(sDate))
            NumDays = DateDiff("d", DayZero, Today)
            DateIconv = CStr(NumDays)
        End Select
    End Select
    Err.Clear
End Function
Public Sub LstViewToWorksheetGroupBy(frmForm As Form, LstView As ListView, ByVal strFile As String, ByVal colName As String, Optional ByVal StrHeader As String = "", Optional ByVal StrLeftFooter As String = "", Optional ByVal StrCenterFooter As String = "", Optional ByVal pOrientation As String = "Landscape", Optional ShowProcess As Boolean = True)
    On Error Resume Next
    Err.Clear
    Exit Sub
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim ColumnNames As String
    Dim spColumns() As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xPageSetUp As Excel.PageSetup
    Dim bExist As Boolean
    Dim sheetName As String
    Dim db As DAO.Database
    Dim Rs As DAO.Recordset
    Dim spLine() As String
    Dim colCnt As Long
    Dim colTot As Long
    Dim cPos As Long
    Dim colData As String
    Dim sFontName As String
    Dim sFontSize As Double
    Dim sheetTot As Long
    Dim sheetCnt As Long
    Dim spSheets() As String
    Dim bFontItalic As Boolean
    strFile = FileName_Validate(strFile)
    bExist = FileExists(strFile)
    If bExist = True Then
        Kill strFile
    End If
    If DirExists(FileToken(strFile, "p")) = False Then
        MakeDirectory FileToken(strFile, "p")
    End If
    If ShowProcess = True Then
        StatusMessage frmForm, "Opening MS Excel..."
    End If
    Set xlApp = New Excel.Application
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.WindowState = Excel.xlMaximized
    xlApp.Visible = False
    ' find the position of the column in the listview
    cPos = LstViewColumnPosition(LstView, colName)
    If ShowProcess = True Then
        StatusMessage frmForm, "Reading the contents of " & colName & ", please be patient..."
    End If
    ' get the contents of the column in question
    colData = LstViewColMV(LstView, cPos, VM, True)
    ' add a new workbook
    xlApp.Workbooks.Add
    Set xlBook = xlApp.ActiveWorkbook
    If ShowProcess = True Then
        StatusMessage frmForm, "Erasing existing work sheets, please be patient..."
    End If
    ' remove currently existing worksheets
    rsTot = xlBook.Sheets.Count
    For rsCnt = rsTot To 2 Step -1
        xlBook.Sheets(rsCnt).Delete
    Next
    ' get the font names and size used
    sFontName = LstView.Font.Name
    sFontSize = LstView.Font.Size
    sFontSize = Round(Val(sFontSize))
    bFontItalic = LstView.Font.Italic
    ' get the column names
    ColumnNames = LstViewColNames(LstView)
    colTot = StrParse(spColumns, ColumnNames, ",")
    ' get the column names and create relevant spreadsheets
    sheetTot = StrParse(spSheets, colData, VM)
    ' create the necessary worksheets
    For sheetCnt = 1 To sheetTot
        sheetName = spSheets(sheetCnt)
        If ShowProcess = True Then
            StatusMessage frmForm, "Creating worksheet " & sheetCnt & "/" & sheetTot & ", please wait..."
        End If
        Set xlSheet = xlBook.Worksheets.Add()
        With xlSheet
            .Activate
            .Name = ExcelCorrectSheetName(sheetName)
            .Cells.NumberFormat = "General"
            .Cells.Font.Name = sFontName
            .Cells.Font.Size = sFontSize
            For colCnt = 1 To colTot
                .Cells(1, colCnt).Value = spColumns(colCnt)
            Next
        End With
    Next
    ' delete sheet 1
    Set xlSheet = xlBook.Worksheets("Sheet1")
    If TypeName(xlSheet) <> "Nothing" Then xlSheet.Delete
    ' save the workbook
    xlBook.SaveAs strFile
    ' Close the Workbook
    xlBook.Close
    ' Close Microsoft Excel with the Quit method.
    xlApp.Quit
    ' open the workbook and treat it as an access database
    Set db = DAO.OpenDatabase(strFile, False, False, "Excel 8.0; HDR=YES;")
    ' loop through each record in lstview and update sheets
    rsTot = LstView.ListItems.Count
    If ShowProcess = True Then
        ProgBarInit frmForm.progBar, rsTot
        StatusMessage frmForm, "Generating reports, please be patient..."
    End If
    For rsCnt = 1 To rsTot
        If ShowProcess = True Then frmForm.progBar.Value = rsCnt
        ' get the contents of the row
        spLine = LstViewGetRow(LstView, rsCnt)
        If LCase$(spLine(1)) = "totals" Then GoTo NextLine
        ' read the sheet name from the control fld
        sheetName = ExcelCorrectSheetName(spLine(cPos))
        Set Rs = db.OpenRecordset(sheetName & "$")
        Rs.AddNew
        For colCnt = 1 To colTot
            Rs(colCnt - 1) = spLine(colCnt)
        Next
        Rs.Update
        Rs.Close
NextLine:
        DoEvents
    Next
    ' count number of records per worksheet
    For sheetCnt = 1 To sheetTot
        sheetName = ExcelCorrectSheetName(spSheets(sheetCnt))
        Set Rs = db.OpenRecordset(sheetName & "$")
        rsTot = Rs.RecordCount
        Rs.Close
        SaveReg sheetName, CStr(rsTot), "records", "LstViewGroupBy"
    Next
    db.Close
    If ShowProcess = True Then
        StatusMessage frmForm
    End If
    ' reopen excel and format report
    If ShowProcess = True Then
        StatusMessage frmForm, "Opening MS Excel..."
    End If
    ' open the database, we need to find out how many records we have per worksheet
    Set xlApp = New Excel.Application
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.WindowState = Excel.xlMaximized
    xlApp.Visible = False
    xlApp.Workbooks.Open strFile
    Set xlBook = xlApp.ActiveWorkbook
    If ShowProcess = True Then
        ProgBarInit frmForm.progBar, sheetTot
    End If
    For sheetCnt = 1 To sheetTot
        sheetName = ExcelCorrectSheetName(spSheets(sheetCnt))
        If ShowProcess = True Then
            frmForm.progBar.Value = sheetCnt
        End If
        Set xlSheet = xlBook.Sheets(sheetName)
        With xlSheet
            .Activate
            For colCnt = 1 To colTot
                .Cells(1, colCnt).Interior.ColorIndex = 15
                .Cells(1, colCnt).Font.Bold = True
                .Cells(1, colCnt).Borders.Weight = Excel.XlBorderWeight.xlThin
                .Cells(1, colCnt).Interior.Pattern = Excel.xlSolid
                .Columns(colCnt).AutoFit
                If LstView.ColumnHeaders(colCnt).Alignment = lvwColumnCenter Then
                    .Columns(colCnt).HorizontalAlignment = Excel.xlCenter
                ElseIf LstView.ColumnHeaders(colCnt).Alignment = lvwColumnRight Then
                    .Columns(colCnt).HorizontalAlignment = Excel.xlRight
                End If
            Next
            rsTot = Val(ReadReg(sheetName, "records", "LstViewGroupBy")) + 1
            .Cells(1, 1).Resize(rsTot, colTot).Borders.Weight = Excel.XlBorderWeight.xlThin
            .Cells(1, 1).Resize(rsTot, colTot).Font.Italic = bFontItalic
            .Cells(1, 1).Resize(rsTot, colTot).VerticalAlignment = Excel.xlTop
            .Rows.AutoFit
        End With
        StrCenterFooter = ProperCase(spSheets(sheetCnt))
        GoSub DoPageSetup
    Next
    xlBook.Save
    ' Close the Workbook
    xlBook.Close
    ' Close Microsoft Excel with the Quit method.
    xlApp.Quit
    ' Release the objects.
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    If ShowProcess = True Then
        StatusMessage frmForm
    End If
    Err.Clear
    Exit Sub
DoPageSetup:
    Set xPageSetUp = xlSheet.PageSetup
    With xPageSetUp
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
        .PrintArea = ""
        .CenterHeader = ""
        .LeftHeader = "&B&I&" & Quote & "Tahoma" & Quote & "&10" & StrHeader
        .RightHeader = ""
        .LeftFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & StrLeftFooter
        .CenterFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & StrCenterFooter
        .RightFooter = "&" & Quote & "Tahoma" & Quote & "&8" & "Page &P of &N"
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintNotes = False
        .CenterHorizontally = False
        .CenterVertically = False
        Select Case LCase$(pOrientation)
        Case "landscape"
            .Orientation = Excel.xlLandscape
        Case Else
            .Orientation = Excel.xlPortrait
        End Select
        .Draft = False
        .PaperSize = Excel.xlPaperA4
        .FirstPageNumber = Excel.xlAutomatic
        .Order = Excel.xlDownThenOver
        .BlackAndWhite = False
        DoEvents
    End With
    Err.Clear
    Return
    Err.Clear
End Sub
Function MvRemoveBlanks(ByVal StrValue As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    xRslt = ""
    Call StrParse(xData, StrValue, Delim)
    xTot = UBound(xData)
    For xCnt = 1 To xTot
        If Len(Trim$(xData(xCnt))) > 0 Then
            xRslt = StringsConcat(xRslt, xData(xCnt), Delim)
        End If
    Next
    xRslt = RemoveDelim(xRslt, Delim)
    MvRemoveBlanks = xRslt
    Err.Clear
End Function
Public Function MvReplaceItem(ByVal StrValue As String, ByVal strItem As String, ByVal StrReplaceWith As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spTot As Long
    Dim spCnt As Long
    Call StrParse(spItems, StrValue, Delim)
    spTot = UBound(spItems)
    For spCnt = 1 To spTot
        If LCase$(spItems(spCnt)) = LCase$(strItem) Then
            spItems(spCnt) = StrReplaceWith
        End If
    Next
    MvReplaceItem = MvFromArray(spItems, Delim)
    Err.Clear
End Function
Public Function ToDate(ByVal Strthedate As String) As String
    On Error Resume Next
    ToDate = Format$(Strthedate, "dd/mm/yyyy")
    Err.Clear
End Function
Public Sub SaveReg(ByVal sKey As String, ByVal sValue As String, Optional ByVal sSection As String = "account", Optional ByVal sAppName As String = "")
    On Error Resume Next
    If Len(sAppName) = 0 Then
        sAppName = App.Title
    End If
    sValue = Replace$(sValue, vbCr, vbNullChar)
    sValue = Replace$(sValue, vbLf, vbNullChar)
    sSection = Replace$(sSection, VM, "\")
    SaveSetting sAppName, sSection, sKey, sValue
    Err.Clear
End Sub
Public Function ReadReg(ByVal sKey As String, Optional ByVal sSection As String = "account", Optional ByVal sAppName As String = "", Optional Default As String = "") As String
    On Error Resume Next
    If Len(sAppName) = 0 Then
        sAppName = App.Title
    End If
    sSection = Replace$(sSection, VM, "\")
    ReadReg = GetSetting(sAppName, sSection, sKey)
    If ReadReg = "" Then
        If Len(Default) > 0 Then
            ReadReg = Default
        End If
    End If
    Err.Clear
End Function
Public Sub LstViewResizeMax(LstView As ListView)
    On Error Resume Next
    Dim col2adjust As Long
    col2adjust = LstView.ColumnHeaders.Count - 1
    Call SendMessage(LstView.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Err.Clear
End Sub
Public Sub LstBoxUpdate(lstBox As Variant, ParamArray items())
    On Error Resume Next
    Dim Item As Variant
    For Each Item In items
        If LstBoxFindExactItemAPI(lstBox, CStr(Item)) = -1 Then
            lstBox.AddItem CStr(Item)
        End If
    Next
    Set Item = Nothing
    Err.Clear
End Sub
Sub AddStatusBar(Sobject As Variant, progBar As Variant)
    On Error Resume Next
    Dim pnlA As Panel
    Dim RowCounter As Integer
    For RowCounter = 1 To 6
        Set pnlA = Sobject.Panels.Add()
    Next
    ' set the style of each panel
    With Sobject.Panels
        .Item(1).Style = 0
        .Item(1).Width = 4005
        .Item(1).Bevel = sbrInset
        .Item(2).Style = 0
        .Item(2).Width = 3000
        .Item(2).Bevel = sbrInset
        .Item(3).Style = 0
        .Item(3).Width = 3000
        .Item(3).Bevel = sbrInset
        .Item(4).Style = 0
        .Item(4).Width = 5000
        .Item(4).Bevel = sbrInset
        .Item(5).Width = 1000
        .Item(5).Bevel = sbrInset
    End With
    Sobject.Refresh
    PutProgressBarInStatusBar Sobject, progBar, 5
    Set pnlA = Nothing
    Err.Clear
End Sub
Sub ResizeStatusBar(objForm As Form, objStatusBar As StatusBar, progBar As ProgressBar)
    On Error Resume Next
    Dim lngSum As Long
    With objStatusBar.Panels
        lngSum = .Item(1).Width + .Item(2).Width + .Item(3).Width + .Item(4).Width
        .Item(5).Width = objForm.Width - lngSum
    End With
    objStatusBar.Refresh
    PutProgressBarInStatusBar objStatusBar, progBar, 5
    Err.Clear
End Sub
Public Function RemAllNL(ByVal StrString As String) As String
    On Error Resume Next
    Dim StrSize As Long
    Dim LAST2 As String
    Dim tmpstring As String
    Dim NL As String
    NL = NL = Chr$(13) + Chr$(10)
    tmpstring = StrString
    LAST2 = Right$(tmpstring, 2)
    Do While LAST2 = NL
        StrSize = Len(tmpstring) - 2
        tmpstring = Left$(tmpstring, StrSize)
        LAST2 = Right$(tmpstring, 2)
    Loop
    RemAllNL = tmpstring
    Err.Clear
End Function
Public Function LstViewUpdate(Arrfields() As String, LstView As ListView, Optional ByVal lstIndex As String = "") As Long
    On Error Resume Next
    Dim ItmX As ListItem
    Dim fldCnt As Integer
    Dim sStr As String
    Dim wCnt As Integer
    Select Case Val(lstIndex)
    Case 0
        Set ItmX = LstView.ListItems.Add()
    Case Else
        Set ItmX = LstView.ListItems(Val(lstIndex))
    End Select
    wCnt = UBound(Arrfields) - 1
    With ItmX
        .Text = Arrfields(1)
        For fldCnt = 1 To wCnt
            .SubItems(fldCnt) = Arrfields(fldCnt + 1)
        Next
    End With
    LstViewUpdate = ItmX.Index
    Set ItmX = Nothing
    'Err.Clear
    Err.Clear
End Function
Function MyAssistant(ByVal StrMsg As String, ByVal strButton As String, ByVal StrIcon As String, ByVal StrHeading As String, Optional ByVal strMode As String = "ma", Optional ByVal strBallonType As String = "but", Optional LabelArray As Variant = Nothing, Optional CheckBoxArray As Variant = Nothing) As Variant
    On Error Resume Next
    ' button can be any of
    ' ync - yesnocancel, c - cancel, o - ok, oc - okcancel, rc - retrycancel and yn - yesno
    ' and ari - abortretryignore, bc - backclose, bnc - backnextclose
    ' bns - backnextsnooze, nc - nextclose, sc - searchclose, toc - tipsoptionsclose, yanc - yesallnocancel
    ' icon can be any of
    ' i - information, w - warning, c - critical, t - tip, q - query
    ' mode can be any of
    ' ad - autodown, ma - modal, me - modeless
    Dim lblCnt As Long
    Dim lblTot As Long
    Dim lblStr As String
    Dim chkSel As String
    Dim isCheck As Long
    Dim objdoc As Word.Document
    Set objdoc = New Word.Document
    isCheck = 0
    With objdoc.Application.Assistant
        If .On = False Then
            .On = True
        End If
        If .Visible = False Then
            .Visible = True
        End If
        .MoveWhenInTheWay = True
        ' create a new ballon
        With .NewBalloon
            Select Case LCase$(strMode)
            Case "ad"
                .Mode = 1           ' msoModeAutoDown
            Case "ma"
                .Mode = 0           ' msoModeModal
            Case "me"
                .Mode = 2           ' msoModeModeless
            End Select
            .Heading = StrHeading
            .Animation = 5          ' msoAnimationRestPose
            Select Case LCase$(strBallonType)
            Case "n"
                .BalloonType = 2    ' msoBalloonTypeNumbers
            Case "bul"
                .BalloonType = 1    ' msoBalloonTypeBullets
            Case "but"
                .BalloonType = 0    ' msoBalloonTypeButtons
            End Select
            Select Case LCase$(strButton)
            Case "ync"
                .Button = 5     'msoButtonSetYesNoCancel
            Case "c"
                .Button = 2     'msoButtonSetCancel
            Case "o"
                .Button = 1     'msoButtonSetOK
            Case "oc"
                .Button = 3     'msoButtonSetOkCancel
            Case "rc"
                .Button = 9     'msoButtonSetRetryCancel
            Case "yn"
                .Button = 4     'msoButtonSetYesNo
            Case "ari"
                .Button = 10    'msoButtonSetAbortRetryIgnore
            Case "bc"
                .Button = 6     'msoButtonSetBackClose
            Case "bnc"
                .Button = 8     'msoButtonSetBackNextClose
            Case "bns"
                .Button = 12    'msoButtonSetBackNextSnooze
            Case "nc"
                .Button = 7     'msoButtonSetNextClose
            Case "sc"
                .Button = 11    'msoButtonSetSearchClose
            Case "toc"
                .Button = 13    'msoButtonSetTipsOptionsClose
            Case "yanc"
                .Button = 14    'msoButtonSetYesAllNoCancel
            Case "none", "n"
                .Button = 0     'msoButtonSetNone
            End Select
            Select Case LCase$(StrIcon)
            Case "i"
                .Icon = 4   'msoIconAlertInfo
            Case "w", "e"
                .Icon = 5   'msoIconAlertWarning
            Case "c"
                .Icon = 7   'msoIconAlertCritical
            Case "t"
                .Icon = 3   'msoIconTip
            Case "q"
                .Icon = 6   'msoAlertQuery
            End Select
            .Text = StrMsg
            ' configure labels
            If IsMissing(LabelArray) = False Then
                lblTot = UBound(LabelArray)
                If lblTot > 5 Then
                    lblTot = 5
                End If
                For lblCnt = 1 To lblTot
                    lblStr = LabelArray(lblCnt)
                    .Labels(lblCnt).Text = lblStr
                Next
            End If
            ' configure checkboxes
            If IsMissing(CheckBoxArray) = False Then
                lblTot = UBound(CheckBoxArray)
                If lblTot > 5 Then
                    lblTot = 5
                End If
                For lblCnt = 1 To lblTot
                    lblStr = CheckBoxArray(lblCnt)
                    isCheck = isCheck + 1
                    .Checkboxes(lblCnt).Text = lblStr
                Next
            End If
            MyAssistant = .Show
            ' find selected indexes of checkboxes
            Select Case isCheck
            Case 0
            Case Else
                chkSel = ""
                For lblCnt = 1 To lblTot
                    If .Checkboxes(lblCnt).Checked = True Then
                        If chkSel = "" Then
                            chkSel = CStr(lblCnt)
                        Else
                            chkSel = chkSel & "," & CStr(lblCnt)
                        End If
                    End If
                Next
                MyAssistant = chkSel
            End Select
        End With
    End With
    MyAssistant = VbAssist(MyAssistant)
    Set objdoc = Nothing
    Err.Clear
End Function
Function VbAssist(varResult As Variant) As Variant
    On Error Resume Next
    Select Case varResult
    Case -1: VbAssist = vbOK
    Case -2: VbAssist = vbCancel
    Case -3: VbAssist = vbYes
    Case -4: VbAssist = vbNo
    Case -5: VbAssist = vbBack
    Case -6: VbAssist = vbNext
    Case -7: VbAssist = vbRetry
    Case -8: VbAssist = vbAbort
    Case -9: VbAssist = vbIgnore
    Case -10: VbAssist = vbSearch
    Case -11: VbAssist = vbSnooze
    Case -12: VbAssist = vbClose
    Case -13: VbAssist = vbTips
    Case -14: VbAssist = vbOptions
    Case -15: VbAssist = vbYesToAll
    Case Else
        VbAssist = varResult
    End Select
    Err.Clear
End Function
Public Sub LstViewCheckAll(LstView As ListView, Optional ByVal bOp As Boolean = True)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = LstView.ListItems.Count
    For lstCnt = 1 To lstTot
        LstView.ListItems(lstCnt).Checked = bOp
    Next
    Err.Clear
End Sub
Public Function FixMaskDate(sDate As Variant) As String
    On Error Resume Next
    DateError = 1
    Select Case sDate.Text
    Case "__/__/____"
        FixMaskDate = sDate.Text
        DateError = 0
    Err.Clear
        Exit Function
    End Select
    Select Case IsDate(sDate.Text)
    Case True
        FixMaskDate = Format$(sDate.Text, "dd/mm/yyyy")
        DateError = 0
    Case Else
        MyPrompt "The value '" & sDate.Text & "' you have entered cannot be" & vbCr & "converted into a valid date. Please enter another value.", "o", "w", "Date Error"
        DateError = 1
        FixMaskDate = sDate.Text
    Err.Clear
        Exit Function
    End Select
    Err.Clear
End Function
Public Sub LstViewRemoveChecked(LstView As ListView, Optional ByVal bCheckedStatus As Boolean = True)
    On Error Resume Next
    Dim bOp As Boolean
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = LstView.ListItems.Count
    For lstCnt = lstTot To 1 Step -1
        bOp = LstView.ListItems(lstCnt).Checked
        If bOp = bCheckedStatus Then
            LstView.ListItems.Remove lstCnt
        End If
    Next
    Err.Clear
End Sub
Public Sub LstViewMakeHeadings(LstView As ListView, ByVal strHeads As String, Optional ByVal ClearItems As String = "")
    On Error Resume Next
    Dim fldCnt As Integer
    Dim FldHead() As String
    Dim fldTot As Integer
    Dim colX As Variant
    Dim cPos As Long
    Call StrParse(FldHead, strHeads, ",")
    fldTot = UBound(FldHead)
    LstView.ColumnHeaders.Clear
    If Len(ClearItems) = 0 Then
        LstView.ListItems.Clear
    End If
    LstView.Sorted = False
    ' first column should be left aligned
    Set colX = LstView.ColumnHeaders.Add(, , ProperCase(FldHead(1)), 1440)
    For fldCnt = 2 To fldTot
        With LstView.ColumnHeaders
            FldHead(fldCnt) = ProperCase(FldHead(fldCnt))
            cPos = ArraySearch(ViewHeadings, FldHead(fldCnt))
            Select Case cPos
            Case 0
                Set colX = .Add(, , FldHead(fldCnt), 1440)
            Case Else
                Set colX = .Add(, , FldHead(fldCnt), 1440, vbRightJustify)
            End Select
        End With
    Next
    LstView.View = lvwReport
    LstView.Checkboxes = True
    LstView.GridLines = True
    LstView.FullRowSelect = True
    LstView.Refresh
    Err.Clear
End Sub
Sub IniHeadings()
    On Error Resume Next
    ReDim ViewHeadings(86)
    ViewHeadings(1) = "Amount"
    ViewHeadings(2) = "Amt"
    ViewHeadings(3) = "Receipt"
    ViewHeadings(4) = "Pay"
    ViewHeadings(5) = "January"
    ViewHeadings(6) = "February"
    ViewHeadings(7) = "March"
    ViewHeadings(8) = "April"
    ViewHeadings(9) = "May"
    ViewHeadings(10) = "June"
    ViewHeadings(11) = "July"
    ViewHeadings(12) = "August"
    ViewHeadings(13) = "September"
    ViewHeadings(14) = "October"
    ViewHeadings(15) = "November"
    ViewHeadings(16) = "December"
    ViewHeadings(17) = "Tot"
    ViewHeadings(18) = "Members"
    ViewHeadings(19) = "Supposed"
    ViewHeadings(20) = "Actual"
    ViewHeadings(21) = "Active"
    ViewHeadings(22) = "Deceased"
    ViewHeadings(23) = "Good"
    ViewHeadings(24) = "Bad"
    ViewHeadings(25) = "Married"
    ViewHeadings(26) = "Divorced"
    ViewHeadings(27) = "Single"
    ViewHeadings(28) = "Widowed"
    ViewHeadings(29) = "Amount"
    ViewHeadings(30) = "Registered"
    ViewHeadings(31) = "Temporal"
    ViewHeadings(32) = "Suspended"
    ViewHeadings(33) = "Size"
    ViewHeadings(34) = "Receipt"
    ViewHeadings(35) = "Payee"
    ViewHeadings(36) = "Paypoint"
    ViewHeadings(37) = "Tax"
    ViewHeadings(38) = "Sales"
    ViewHeadings(39) = "Cash"
    ViewHeadings(40) = "Debit"
    ViewHeadings(41) = "Credit"
    ViewHeadings(42) = "Exclusive"
    ViewHeadings(43) = "Inclusive"
    ViewHeadings(44) = "Start"
    ViewHeadings(45) = "End"
    ViewHeadings(46) = "Current"
    ViewHeadings(47) = "Difference"
    ViewHeadings(48) = "Number"
    ViewHeadings(49) = "Membership Number"
    ViewHeadings(50) = "Member #"
    ViewHeadings(51) = "Membership #"
    ViewHeadings(52) = "Age"
    ViewHeadings(53) = "Average"
    ViewHeadings(54) = "Premium"
    ViewHeadings(55) = "Regions"
    ViewHeadings(56) = "Id"
    ViewHeadings(57) = "Id Number"
    ViewHeadings(58) = "Id #"
    ViewHeadings(59) = "Id No."
    ViewHeadings(60) = "Id No"
    ViewHeadings(61) = "Incomes"
    ViewHeadings(62) = "Expenses"
    ViewHeadings(63) = "Totals"
    ViewHeadings(64) = "Real Receipts"
    ViewHeadings(65) = "Actual Receipts"
    ViewHeadings(66) = "Supposed Receipts"
    ViewHeadings(67) = "Qty"
    ViewHeadings(68) = "Quantity"
    ViewHeadings(69) = "Cash Sales"
    ViewHeadings(70) = "Member Sales"
    ViewHeadings(71) = "Sales"
    ViewHeadings(72) = "Price"
    ViewHeadings(73) = "Discount"
    ViewHeadings(74) = "Tax %"
    ViewHeadings(75) = "Disc %"
    ViewHeadings(76) = "Tax Amount"
    ViewHeadings(77) = "Discount Amount"
    ViewHeadings(78) = "Target"
    ViewHeadings(79) = "Granted"
    ViewHeadings(80) = "Recorded Expenses"
    ViewHeadings(81) = "Recorded Difference"
    ViewHeadings(82) = "Amount Required"
    ViewHeadings(83) = "Amount Received"
    ViewHeadings(84) = "Required"
    ViewHeadings(85) = "Provided"
    ViewHeadings(86) = "Variance"
    Err.Clear
End Sub
Public Sub TreeViewClearAPI(TreeV As TreeView)
    On Error Resume Next
    Dim lNodeHandle As Long
    Dim tvHwnd As Long
    tvHwnd = TreeV.hWnd
    ' Turn off redrawing on the Treeview for more speed improvements
    Do
        lNodeHandle = SendMessageLONG(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        If lNodeHandle > 0 Then
            SendMessageLONG tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
        Else
            Exit Do
        End If
    Loop
    Err.Clear
End Sub
Function TreeViewAddPath(TreeV As TreeView, ByVal sPath As String, Optional ByVal Image As String = "", Optional ByVal SelectedImage As String = "", Optional ByVal Tag As String = "", Optional Delim As String = "\") As Long
    On Error Resume Next
    Dim prevP As String
    Dim currP As String
    Dim lngC As Long
    Dim lngT As Long
    Dim pStr() As String
    Dim currN As String
    Dim nodeN As Node
    Dim pKey As String
    Dim cLoc As Long
    Call StrParse(pStr, sPath, Delim)
    lngT = UBound(pStr)
    For lngC = 1 To lngT
        prevP = MvFromMv(sPath, 1, lngC - 1, Delim)
        currP = MvFromMv(sPath, 1, lngC, Delim)
        currN = pStr(lngC)
        If prevP = "" Then
            ' this is the root node
            cLoc = TreeViewSearchPath(TreeV, currP)
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(, , currP, currP)
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        Else
            ' this is the second, third etc node
            cLoc = TreeViewSearchPath(TreeV, currP)
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(pKey, tvwChild, currP, currN)
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        End If
        pKey = nodeN.Key
        If lngC = lngT Then
            TreeViewAddPath = nodeN.Index
            Exit For
        End If
    Next
    Err.Clear
End Function
Function TreeViewPathLocation(treeDms As TreeView, ByVal SearchPath As String) As Long
    On Error Resume Next
    Dim myNode As Node
    TreeViewPathLocation = 0
    For Each myNode In treeDms.Nodes
        If LCase$(myNode.FullPath) = LCase$(SearchPath) Then
            TreeViewPathLocation = myNode.Index
            Exit For
        End If
    Next
    Err.Clear
End Function
Public Function DialogOpen(CD As CommonDialog, Optional ByVal Title As String = "Open Existing File", Optional ByVal InitDir As String = "...", Optional ByVal DefaultExt As String = "*.*") As String
    On Error GoTo ErrHandler
    Dim filName As String
    Dim filCnt As Integer
    Dim spFilter() As String
    Dim strFilter As String
    strFilter = FileFilters
    filCnt = 0
    Call StrParse(spFilter, strFilter, "|")
    filCnt = ArraySearch(spFilter, DefaultExt)
    If filCnt <> 0 Then
        filCnt = filCnt / 2
    End If
    filName = ""
    With CD
        .CancelError = True
        '.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        .Filter = strFilter
        .DialogTitle = Title
        .InitDir = InitDir
        .Filename = ""
        .DefaultExt = DefaultExt
        .FilterIndex = filCnt
        .ShowOpen
        filName = .Filename
        If Len(filName) = 0 Then
    Err.Clear
            Exit Function
        End If
    End With
    DialogOpen = filName
    SaveReg "lastpath", FileToken(filName, "p")
    Err.Clear
    Exit Function
ErrHandler:
    Err.Clear
    Exit Function
    Err.Clear
End Function
Function FileFilters() As String
    On Error Resume Next
    Dim s_result As String
    s_result = "All Files (*.*)|*.*|Template File (*.tem)|*.tem"
    s_result = s_result & "|Temporal File (*.tmp)|*.tmp|Transaction File (*.trn)|*.trn"
    s_result = s_result & "|Data File (*.dat)|*.dat|Settings File (*.ini)|*.ini"
    s_result = s_result & "|Wave File (*.wav)|*.wav|Mpeg 3 File (*.mp3)|*.mp3"
    s_result = s_result & "|Help File Creator (*.hfc)|*.hfc|Bible File (*.bib)|*.bib"
    s_result = s_result & "|Dictionary File (*.dic)|*.dic|Topic Note File (*.top)|*.top"
    s_result = s_result & "|Study Node File (*.stu)|*.stu|Commentary File (*.com)|*.com"
    s_result = s_result & "|Graphics File (*.gra)|*.gra|Audio File (*.aud)|*.aud"
    s_result = s_result & "|Comma-Separated Values (*.csv)|*.csv|Sequential Access (*.seq)|*.seq"
    s_result = s_result & "|Excel (*.xls)|*.xls|Lotus 123 (*.wks)|*.wks"
    s_result = s_result & "|Rich Text Format (*.rtf)|*.rtf|Text (*.txt)|*.txt"
    s_result = s_result & "|Word for Windows (*.doc)|*.doc|Microsoft Access (*.mdb)|*.mdb"
    s_result = s_result & "|Adobe Acrobat (*.pdf)|*.pdf"
    s_result = s_result & "|Tree File (*.tree)|*.tree|Dictionary File (*.dict)|*.dict"
    s_result = s_result & "|Visual Basic Project File (*.vbp)|*.vbp|Visual Basic Project Group File (*.vbg)|*.vbg"
    s_result = s_result & "|Visual Basic Mak File (*.mak)|*.mak|Visual Basic Form File(*.frm)|*.frm"
    s_result = s_result & "|Visual Basic Module File (*.mod)|*.mod|Visual Basic Class Module (*.cls)|*.cls"
    s_result = s_result & "|Bitmap File (*.bmp)|*.bmp|Tif File (*.tif)|*.tif"
    s_result = s_result & "|Tiff File (*.tif)|*.tif|Jpeg File (*.jpg)|*.jpg"
    s_result = s_result & "|Gif File (*.gif)|*.gif|Png File (*.png)|*.png"
    s_result = s_result & "|Batch File (*.bat)|*.bat|Executable File (*.exe)|*.exe"
    s_result = s_result & "|Icon File (*.ico|*.ico|Configuration (*.cfg)|*.cfg"
    s_result = s_result & "|Visual Basic Setup File (*.lst)|*.lst|Inno Setup Script (*.iss)|*.iss"
    s_result = s_result & "|Spell Check Log (*.spl)|*.spl|Document Tracking System (*.dts)|*.dts"
    s_result = s_result & "|Archive File (*.arc)|*.arc|List View File (*.lvf)|*.lvf"
    s_result = s_result & "|Executable File (*.exe)|*.exe|Dynamic Link Library File (*.dll)|*.dll"
    FileFilters = s_result
    Err.Clear
End Function
Public Function NextNewFile(ByVal StrFilePath As String, Optional IsCopy As Boolean = False) As String
    On Error Resume Next
    Dim bExist As Boolean
    Dim fCnt As Long
    Dim sExtension As String
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    Dim pPath As String
    Dim fName As String
    Set fsoFile = fso.GetFile(StrFilePath)
    pPath = fso.GetParentFolderName(StrFilePath)
    If Right$(pPath, 1) = "\" Then
        pPath = Left$(pPath, Len(pPath) - 1)
    End If
    fName = fso.GetBaseName(StrFilePath)
    sExtension = fso.GetExtensionName(StrFilePath)
    fCnt = 0
    bExist = IsPathFile(StrFilePath)
    Do Until bExist = False
        fCnt = fCnt + 1
        StrFilePath = pPath & "\" & fName & " " & CStr(fCnt) & "." & sExtension
        If IsCopy = True Then
            If fCnt - 1 < 0 Then
                StrFilePath = pPath & "\Copy Of " & fName & "." & sExtension
            Else
                StrFilePath = pPath & "\Copy " & CStr(fCnt) & " Of " & fName & "." & sExtension
            End If
        End If
        DoEvents
        bExist = IsPathFile(StrFilePath)
    Loop
    NextNewFile = StrFilePath
    Err.Clear
End Function
Public Function IsPathFile(ByVal strPath As String) As Boolean
    On Error Resume Next
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    If fso.FileExists(strPath) = True Then
        Set fsoFile = fso.GetFile(strPath)
        If TypeName(fsoFile) = "Nothing" Then
            IsPathFile = False
        Else
            IsPathFile = True
        End If
    Else
        IsPathFile = False
    End If
    Set fsoFile = Nothing
    Set fso = Nothing
    Err.Clear
End Function
Function FormatTextForWord(ByVal StrValue As String, Optional ByVal NumberOfTabs As Long = 4) As String
    On Error Resume Next
    Dim spTot As Long
    Dim spCnt As Long
    Dim spDat() As String
    Dim sTabs As String
    Dim NL As String
    NL = NL = Chr$(13) + Chr$(10)
    sTabs = String$(NumberOfTabs, vbTab)
    Call StrParse(spDat, StrValue, NL)
    spTot = UBound(spDat)
    For spCnt = 2 To spTot
        spDat(spCnt) = sTabs & spDat(spCnt)
    Next
    FormatTextForWord = MvFromArray(spDat, NL)
    Err.Clear
End Function
Sub ReloadMenu(MenuControl As Variant, LstFrom As Variant, Optional ByVal StrOwn As String = "", Optional ByVal StrBlank As String = "", Optional ByVal Ascending As Boolean = True, Optional ByVal NumberOfItemsToLoad As Long = -1, Optional MakeProperCase As Boolean = False)
    On Error Resume Next
    Dim menuTot As Long
    Dim menuCnt As Long
    Dim menuItm As Long
    Dim cntItem As Long
    Dim menuCnt_Cnt As Long
    cntItem = 0
    menuCnt = 0
    ' how many items are there yet
    menuTot = MenuControl.Count - 1
    For menuCnt = menuTot To 1 Step -1
        Unload MenuControl(menuCnt)
    Next
    menuItm = LstFrom.ListCount - 1
    If Ascending = True Then
        MenuControl(0).Caption = IIf((MakeProperCase = True), ProperCase(LstFrom.List(0)), LstFrom.List(0))
        MenuControl(0).Checked = False
    Else
        MenuControl(0).Caption = IIf((MakeProperCase = True), ProperCase(LstFrom.List(menuItm)), LstFrom.List(menuItm))
        MenuControl(0).Checked = False
    End If
    If MenuControl(0).Caption = "" Then
        MenuControl(0).Caption = "<No Records>"
    End If
    If menuItm = -1 Then
        MenuControl(0).Enabled = False
    Else
        MenuControl(0).Enabled = True
    End If
    Select Case Ascending
    Case True
        If NumberOfItemsToLoad < 0 Then
            NumberOfItemsToLoad = menuItm
        End If
        For menuCnt = 1 To NumberOfItemsToLoad
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = IIf((MakeProperCase = True), ProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
        Next
    Case Else
        If NumberOfItemsToLoad < 0 Then
            menuCnt_Cnt = menuItm - 1
            For menuCnt = menuCnt_Cnt To 1 Step -1
                cntItem = cntItem + 1
                Load MenuControl(cntItem)
                MenuControl(cntItem).Caption = IIf((MakeProperCase = True), ProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
            Next
        Else
            menuCnt_Cnt = menuItm - 1
            For menuCnt = menuCnt_Cnt To 1 Step -1
                cntItem = cntItem + 1
                If cntItem + 1 > NumberOfItemsToLoad Then
                    Exit For
                End If
                Load MenuControl(cntItem)
                MenuControl(cntItem).Caption = IIf((MakeProperCase = True), ProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
            Next
        End If
    End Select
    If Len(StrOwn) > 0 Then
        Select Case MenuControl(0).Caption
        Case "<No Records>"
            MenuControl(0).Enabled = True
            MenuControl(0).Caption = "<Enter Own>"
        Case Else
            menuCnt = MenuControl.Count - 1
            menuCnt = menuCnt + 1
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = "<Enter Own>"
        End Select
    End If
    If Len(StrBlank) > 0 Then
        Select Case MenuControl(0).Caption
        Case "<No Records>"
            MenuControl(0).Enabled = True
            MenuControl(0).Caption = "<Blank>"
        Case Else
            menuCnt = MenuControl.Count - 1
            menuCnt = menuCnt + 1
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = "<Blank>"
        End Select
    End If
    Err.Clear
End Sub
Public Sub LstBoxRemoveItemsAPI(lstBox As Variant, ParamArray items())
    On Error Resume Next
    Dim Item As Variant
    Dim ItemPos As Long
    Dim ItemName As String
    For Each Item In items
        ItemName = CStr(Item)
        ItemPos = LstBoxFindExactItemAPI(lstBox, ItemName)
        If ItemPos <> -1 Then
            lstBox.RemoveItem ItemPos
        End If
    Next
    Set Item = Nothing
    Err.Clear
End Sub
Public Sub ShowFrame(fraObject As Variant, ByVal L As Long, ByVal t As Long, ByVal W As Long, ByVal h As Long, Optional ByVal Caption As String = "", Optional ShowBorder As Boolean = True)
    On Error Resume Next
    With fraObject
        .Left = L
        .Top = t
        .Width = W
        .Height = h
        .Visible = True
        If ShowBorder = False Then .BorderStyle = 0
        If IsMissing(Caption) = False Then .Caption = Caption
        .ZOrder 0
    End With
    Err.Clear
End Sub
Public Function StringToMv(ByVal Delim As String, ParamArray items()) As String
    On Error Resume Next
    Dim Item As Variant
    Dim NewString As String
    Dim NewItem As String
    NewString = ""
    For Each Item In items
        NewItem = CStr(Item)
        NewString = StringAdd(NewString, NewItem, Delim)
    Next
    StringToMv = RemoveDelim(NewString, Delim)
    Err.Clear
End Function
Sub FormatRTF(objVariant As Variant)
    On Error Resume Next
    objVariant.SelStart = 0
    objVariant.SelLength = Len(objVariant.Text)
    objVariant.SelFontName = "Tahoma"
    objVariant.SelFontSize = 8.25
    Err.Clear
End Sub
Public Sub LstViewSumColumns(LstView As ListView, ToMoney As Boolean, ParamArray myColumns())
    On Error Resume Next
    Dim rsCnt As Long
    Dim myColumn As Variant
    Dim colSum As String
    Dim totPos As Long
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spColumns() As String
    totPos = LstViewFindItem(LstView, "Totals", search_Text, search_Whole)
    If totPos = 0 Then
        LstView.ListItems.Add , , "Totals"
    End If
    For Each myColumn In myColumns
        spTot = StrParse(spColumns, CStr(myColumn), ",")
        For spCnt = 1 To spTot
            rsCnt = LstViewColumnPosition(LstView, spColumns(spCnt))
            colSum = LstViewSumColumn(LstView, rsCnt, ToMoney, True)
            totPos = LstViewFindItem(LstView, "Totals", search_Text, search_Whole)
            spLine = LstViewGetRow(LstView, totPos)
            If ToMoney = True Then
                colSum = MakeMoney(colSum)
            Else
                colSum = Format$(colSum, "#,###")
            End If
            spLine(rsCnt) = colSum
            totPos = LstViewUpdate(spLine, LstView, CStr(totPos))
            LstView.ListItems(totPos).EnsureVisible
        Next
    Next
    Err.Clear
End Sub
Public Function LstViewSumColumn(lstReport As ListView, colPos As Long, Optional ToMoney As Boolean = False, Optional RightAlign As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim strSum As String
    strSum = "0"
    rsTot = lstReport.ListItems.Count
    For rsCnt = 1 To rsTot
        spLine = LstViewGetRow(lstReport, rsCnt)
        If spLine(1) = "Totals" Then
        Else
            If ToMoney = True Then
                strSum = Val(ProperAmount(strSum)) + Val(ProperAmount(spLine(colPos)))
                strSum = ProperAmount(strSum)
            Else
                strSum = Val(strSum) + Val(spLine(colPos))
            End If
        End If
        If ToMoney = True Then
            spLine(colPos) = MakeMoney(spLine(colPos))
        Else
            spLine(colPos) = Format$(spLine(colPos), "#,###")
        End If
        Call LstViewUpdate(spLine, lstReport, CStr(rsCnt))
        If RightAlign = True Then
            lstReport.ColumnHeaders(colPos).Alignment = lvwColumnRight
        End If
    Next
    If ToMoney = True Then
        LstViewSumColumn = ProperAmount(strSum)
    Else
        LstViewSumColumn = Format$(strSum, "#,###")
    End If
    lstReport.Refresh
    Err.Clear
End Function
Public Function LstViewFindItem(LstView As ListView, ByVal StrSearch As String, Optional ByVal SearchWhere As FindWhere = search_Text, Optional SearchItemType As SearchType = search_Whole) As Long
    On Error Resume Next
    Dim itmFound As ListItem
    LstViewFindItem = 0
    Set itmFound = LstView.FindItem(StrSearch, SearchWhere, , SearchItemType)
    If TypeName(itmFound) = "Nothing" Then
    Err.Clear
        Exit Function
    End If
    LstViewFindItem = CLng(itmFound.Index)
    Set itmFound = Nothing
    Err.Clear
End Function
Function Oconv(ByVal sValue As String, Optional ByVal sFormat As String = "", Optional ByVal ValueFormat As String = "") As String
    On Error Resume Next
    Dim theDate As String
    Dim spDate() As String
    Dim syy As String
    Dim hash(1 To 2) As String
    Dim sdd As String
    Dim smm As String
    Dim StrSize As Integer
    If Len(sFormat) = 0 Then
        sFormat = "M"
    End If
    Select Case UCase$(sFormat)
    Case "MDY", "DMY", "YDM", "YMD", "DYM", "MYD"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        syy = Right$(spDate(3), 2)
        Select Case UCase$(sFormat)
        Case "MDY":         Oconv = StringsConcat(spDate(2), spDate(1), syy)
        Case "DMY":         Oconv = StringsConcat(spDate(1), spDate(2), syy)
        Case "YMD":         Oconv = StringsConcat(syy, spDate(2), spDate(1))
        Case "YDM":         Oconv = StringsConcat(syy, spDate(1), spDate(2))
        Case "DYM":         Oconv = StringsConcat(spDate(1), syy, spDate(2))
        Case "MYD":         Oconv = StringsConcat(spDate(2), syy, spDate(1))
        End Select
    Case "MDYY", "DMYY", "YYDM", "YYMD", "DYYM", "MYYD"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        Select Case UCase$(sFormat)
        Case "MDYY":         Oconv = StringsConcat(spDate(2), spDate(1), spDate(3))
        Case "DMYY":         Oconv = StringsConcat(spDate(1), spDate(2), spDate(3))
        Case "YYMD":         Oconv = StringsConcat(spDate(3), spDate(2), spDate(1))
        Case "YYDM":         Oconv = StringsConcat(spDate(3), spDate(1), spDate(2))
        Case "DYYM":         Oconv = StringsConcat(spDate(1), spDate(3), spDate(2))
        Case "MYYD":         Oconv = StringsConcat(spDate(2), spDate(3), spDate(1))
        End Select
    Case "YYMM", "YYYYMM"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        Oconv = StringsConcat(spDate(3), spDate(2))
    Case "F"
        Oconv = Replace$(sValue, "%", "/")
    Case "M"
        If sValue = "." Then
            sValue = "000"
        End If
        sValue = DotAmount(sValue)
        Oconv = Format$(sValue, "#,##0.00")
    Case "D"
        StrSize = Len(sValue)
        Select Case StrSize
        Case Is < 6
            Oconv = sValue
    Err.Clear
            Exit Function
        Case 6
            ValueFormat = FixText(ValueFormat)
            If Len(ValueFormat) = 0 Then
                ValueFormat = "DMY"
            End If
            Select Case ValueFormat
            Case "DMY"
                sdd = Left$(sValue, 2)
                smm = Mid$(sValue, 3, 2)
                syy = Right$(sValue, 2)
            Case "YMD"
                sdd = Right$(sValue, 2)
                smm = Mid$(sValue, 3, 2)
                syy = Left$(sValue, 2)
            Case "YDM"
                syy = Left$(sValue, 2)
                sdd = Mid$(sValue, 3, 2)
                smm = Right$(sValue, 2)
            End Select
            Oconv = ToDate(StringToMv("/", sdd, smm, syy))
        Case 8
            hash(1) = Mid$(sValue, 3, 1)
            hash(2) = Mid$(sValue, 6, 1)
            If (hash(1) = "/") And (hash(2) = "/") Then
                Oconv = sValue
    Err.Clear
                Exit Function
            End If
            ValueFormat = FixText(ValueFormat)
            If Len(ValueFormat) = 0 Then
                ValueFormat = "DMY"
            End If
            Select Case ValueFormat
            Case "DMY"
                sdd = Left$(sValue, 2)
                smm = Mid$(sValue, 3, 2)
                syy = Right$(sValue, 4)
            Case "YMD"
                sdd = Right$(sValue, 2)
                smm = Mid$(sValue, 5, 2)
                syy = Left$(sValue, 4)
            Case "YDM"
                syy = Left$(sValue, 4)
                sdd = Mid$(sValue, 5, 2)
                smm = Right$(sValue, 2)
            End Select
            Oconv = ToDate(StringToMv("/", sdd, smm, syy))
        Case Is >= 10
            hash(1) = Mid$(sValue, 3, 1)
            hash(2) = Mid$(sValue, 6, 1)
            If (hash(1) = "/") And (hash(2) = "/") Then
                Oconv = sValue
            End If
        End Select
    End Select
    Err.Clear
End Function
Public Function DateOconv(ByVal sDays As String) As String
    On Error Resume Next
    DateOconv = sDays
    If Len(sDays) = 0 Then
    Err.Clear
        Exit Function
    End If
    Select Case sDays
    Case Is <> ""
        Select Case IsDate(sDays)
        Case False
            Dim DayZero As Date
            Dim Today As Date
            ' for pick and universe date zero is 31/12/1967
            DayZero = ToDate("31/12/1967")
            Today = DateAdd("d", CDbl(sDays), DayZero)
            DateOconv = ToDate(Today)
        End Select
    End Select
    Err.Clear
End Function
Public Function MonthYearDesc(ByVal Yyyymm As String) As String
    On Error Resume Next
    Dim smm As String
    Dim syy As String
    MonthYearDesc = Yyyymm
    If Len(Yyyymm) = 0 Then
    Err.Clear
        Exit Function
    End If
    smm = Right$(Yyyymm, 2)
    syy = lngYearFrom(Yyyymm)
    MonthYearDesc = StringAdd(StrMonthName(smm), " ", syy)
    Err.Clear
End Function
Public Function YearMonthDesc(ByVal Yyyymm As String) As String
    On Error Resume Next
    Dim smm As String
    Dim syy As String
    YearMonthDesc = Yyyymm
    If Len(Yyyymm) = 0 Then
    Err.Clear
        Exit Function
    End If
    smm = Right$(Yyyymm, 2)
    syy = lngYearFrom(Yyyymm)
    YearMonthDesc = StringAdd(syy, " ", StrMonthName(smm))
    Err.Clear
End Function
Public Function MonthDesc(ByVal Yyyymm As String) As String
    On Error Resume Next
    Dim smm As String
    Dim syy As String
    MonthDesc = Iconv(Yyyymm)
    If Len(Yyyymm) = 0 Then
    Err.Clear
        Exit Function
    End If
    smm = Right$(Yyyymm, 2)
    syy = lngYearFrom(Yyyymm)
    MonthDesc = StrMonthName(smm)
    Err.Clear
End Function
Public Function DotAmount(ByVal sAmount As String) As String
    On Error Resume Next
    Dim s_size As Integer
    Dim s_cents As String
    Dim s_numbers As String
    Dim s_firstpart As Integer
    Dim s_fpart As String
    Dim s_epart As String
    Dim StrSize As Integer
    Select Case Trim$(sAmount)
    Case Is <> ""
        Select Case InStr(sAmount, ".")
        Case 0 ' the amount has no dot
            DotAmount = sAmount & ".00"
            StrSize = Len(sAmount)
            Select Case StrSize
            Case 1
                sAmount = sAmount & "00"
            Case 2
                s_fpart = Left$(sAmount, 1)
                s_epart = Right$(sAmount, 1)
                Select Case s_fpart
                Case "-": sAmount = StringAdd(s_fpart, "00", s_epart)
                Case Else: sAmount = sAmount & "00"
                End Select
            End Select
            s_size = Len(sAmount)
            s_cents = Right$(sAmount, 2)   ' the last two values
            s_firstpart = s_size - 2
            s_numbers = Left$(sAmount, s_firstpart)
            DotAmount = StringAdd(s_numbers, ".", s_cents)
        Case Else
            DotAmount = sAmount
        End Select
    Case ""
        DotAmount = "0.00"
    Case Else
        DotAmount = sAmount
    End Select
    Err.Clear
End Function
Public Function FixText(ByVal sString As String) As String
    On Error Resume Next
    FixText = UCase$(Trim$(sString))
    Err.Clear
End Function
Public Function lngYearFrom(ByVal Yyyymm As String) As Long
    On Error Resume Next
    Dim smm As String
    Dim syy As String
    Dim dLen As Integer
    dLen = Len(Yyyymm) - 2
    smm = Right$(Yyyymm, 2)
    syy = Left$(Yyyymm, dLen)
    Select Case smm
    Case Is >= "13"
        smm = "01"
        syy = Val(syy) + 1
    End Select
    lngYearFrom = CLng(Val(syy))
    Err.Clear
End Function
Public Function StrMonthName(ByVal strMonth As String) As String
    On Error Resume Next
    Dim lngMo As Long
    StrMonthName = strMonth
    If Len(strMonth) = 0 Then
    Err.Clear
        Exit Function
    End If
    lngMo = Val(Right$(strMonth, 2))
    Select Case lngMo
    Case 1 To 12
        StrMonthName = MonthName(lngMo)
    End Select
    Err.Clear
End Function
Public Sub ArrayFromCollection(objCollection As Collection, objArray() As String)
    On Error Resume Next
    Dim lngMin As Long
    Dim lngMax As Long
    lngMax = objCollection.Count
    ReDim Preserve objArray(lngMax)
    For lngMin = 1 To lngMax
        objArray(lngMin) = objCollection.Item(lngMin)
    Next
    Err.Clear
End Sub
Function BuildSQL(ByVal sType As String, ByVal SearchField As String, ByVal StrValues As String, Optional ByVal StrOperation As String = "like", Optional ByVal AndOr As String = "and", Optional DontBreakItemToSearch As Boolean = False) As String
    On Error Resume Next
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim rslt As String
    Dim Delim As String
    rslt = ""
    Delim = ","
    If Len(StrOperation) = 0 Then
        StrOperation = "like"
    End If
    If DontBreakItemToSearch = True Then
    Else
        StrValues = Replace$(StrValues, " ", ",")
    End If
    Call StrParse(spLine, StrValues, Delim)
    spTot = UBound(spLine)
    sType = LCase$(sType)
    For spCnt = 1 To spTot
        spStr = Trim$(spLine(spCnt))
        If Len(spStr) = 0 Then GoTo NextResult
        Select Case StrOperation
        Case "="
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '" & spStr & "' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '" & EscIn(spStr) & "' " & AndOr & " "
            End If
        Case "having"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '" & spStr & "' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '" & EscIn(spStr) & "' " & AndOr & " "
            End If
        Case "like"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '*" & spStr & "*' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '%" & EscIn(spStr) & "%' " & AndOr & " "
            End If
        Case "likestart", "startwith"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '" & spStr & "*' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '" & EscIn(spStr) & "%' " & AndOr & " "
            End If
        Case "likeend", "endwith"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '*" & spStr & "' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '%" & EscIn(spStr) & "' " & AndOr & " "
            End If
        End Select
        rslt = StringsConcat(rslt, spStr)
NextResult:
    Next
    rslt = RemoveDelim(rslt, AndOr & " ")
    BuildSQL = rslt
    Err.Clear
End Function
Public Function IsPwdValid(ByVal StrValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, StrValue, "[") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "]") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ".") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "*") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ">") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "<") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ",") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "`") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "#") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "!") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "/") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "\") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "|") > 0), 1, 0)
    intH = intH + IIf(Len(StrValue) < 8, 1, 0)
    If intH = 0 Then
        IsPwdValid = True
    Else
        IsPwdValid = False
    End If
    Err.Clear
End Function
Public Function HasSpecial(ByVal StrValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, StrValue, ".") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "@") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "$") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "%") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "^") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "&") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "(") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ")") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "-") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "}") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "{") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ":") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, ";") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "?") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "~") > 0), 1, 0)
    If intH = 0 Then
        HasSpecial = False
    Else
        HasSpecial = True
    End If
    Err.Clear
End Function
Public Function HasNumber(ByVal StrValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, StrValue, "0") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "1") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "2") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "3") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "4") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "5") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "6") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "7") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "8") > 0), 1, 0)
    intH = intH + IIf((InStr(1, StrValue, "9") > 0), 1, 0)
    If intH = 0 Then
        HasNumber = False
    Else
        HasNumber = True
    End If
    Err.Clear
End Function
Public Function IsBlank(ObjectName As Variant, ByVal fldName As String, Optional MsgType As Integer = 1) As Boolean
    On Error Resume Next
    Dim strM As String
    Dim StrT As String
    Dim strO As String
    Dim strK As String
    IsBlank = False
    If TypeOf ObjectName Is TextBox Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "type"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ComboBox Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ImageCombo Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is CheckBox Then
        If ObjectName.Value = 0 Then
            strO = "select"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ListBox Then
        If (ObjectName.ListCount - 1) = -1 Then
            strO = "select"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is OptionButton Then
        If ObjectName.Value = False Then
            strO = "select"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is Label Then
        If Len(Trim$(ObjectName.Caption)) = 0 Then
            strO = "specify"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "MaskEdBox" Then
        strK = ObjectName.Mask
        strK = Replace$(strK, "#", "_")
        If Trim$(ObjectName.Text = strK) Then
            strO = "enter"
            GoSub CompileError
            If MsgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            Else
                MyAssistant strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    End If
    Err.Clear
    Exit Function
CompileError:
    strM = "The " & LCase$(fldName) & " cannot be left blank. Please " & strO & " the " & LCase$(fldName) & "."
    StrT = ProperCase(fldName & " error")
    Err.Clear
    Return
    Err.Clear
End Function
Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
    On Error Resume Next
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowseForFolder = sPath
    Err.Clear
End Function
Public Sub LstViewFromComputerFolder(lstReport As ListView, ByVal StrFolder As String, ByVal StrHeading As String, Optional ByVal xPattern As String = "*.txt")
    On Error Resume Next
    If Len(StrFolder) = 0 Then Exit Sub
    Dim colFiles As New Collection
    lstReport.View = lvwReport
    lstReport.Sorted = True
    LstViewMakeHeadings lstReport, StrHeading
    Set colFiles = MyFilesCollection(StrFolder, xPattern)
    LstViewFromCollection lstReport, colFiles
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Public Sub LstViewFromFolder(lstReport As ListView, ByVal StrFolder As String, ByVal StrHeading As String, Optional ByVal xPattern As String = "*.txt")
    On Error Resume Next
    If Len(StrFolder) = 0 Then Exit Sub
    Dim colFiles As New Collection
    lstReport.View = lvwReport
    lstReport.Sorted = True
    LstViewMakeHeadings lstReport, StrHeading
    Set colFiles = MyFilesCollection(StrFolder, xPattern)
    LstViewFromCollection lstReport, colFiles
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Function MyFilesCollection(ByVal StrFolder As String, Optional ByVal StrPattern As String = "*.*") As Collection
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim strFile As String
    Dim strP As String
    Dim colNew As New Collection
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFolder As Scripting.Folder
    Dim fsoFile As Scripting.File
    strP = LCase$(MvField(StrPattern, 2, "."))
    Set fsoFolder = fso.GetFolder(StrFolder)
    If TypeName(fsoFolder) = "Nothing" Then Exit Function
    If TypeName(fsoFolder) = "Nothing" Then Exit Function
    rsTot = fsoFolder.Files.Count
    For Each fsoFile In fsoFolder.Files
        strFile = fsoFile.Path
        If StrPattern = "*.*" Then
            colNew.Add strFile
        Else
            If strP = LCase$(Right$(strFile, Len(strP))) Then
                colNew.Add strFile
            End If
        End If
    Next
    Set MyFilesCollection = colNew
    Err.Clear
End Function
Public Sub LstViewFromCollection(LstView As ListView, varCollection As Collection, Optional ByVal Delim As String = "", Optional ByVal lstClear As String = "", Optional ByVal MaxLevel As Long = -1)
    On Error Resume Next
    Dim spLine() As String
    Dim varTot As Long
    Dim varCnt As Long
    Dim xTot As Long
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    If Len(lstClear) = 0 Then LstView.ListItems.Clear
    varTot = varCollection.Count
    LstViewSetMemory LstView, varTot
    For varCnt = 1 To varTot
        Call StrParse(spLine, varCollection.Item(varCnt), Delim)
        If MaxLevel = -1 Then
            Call LstViewUpdate(spLine, LstView)
        Else
            xTot = UBound(spLine)
            If xTot = MaxLevel Then
                Call LstViewUpdate(spLine, LstView)
            End If
        End If
    Next
    Err.Clear
End Sub
Public Function FileData(ByVal Filename As String) As String
    On Error Resume Next
    Dim sLen As Long
    Dim fileNum As Long
    Dim Size As Long
    fileNum = FreeFile
    Size = FileLen(Filename)
    Open Filename For Input Access Read As #fileNum
        sLen = LOF(fileNum)
        FileData = Input(sLen, #fileNum)
    Close #fileNum
    Err.Clear
End Function
Public Function IsUploadable(ByVal StrValue As String) As Boolean
    On Error Resume Next
    Dim fPart As String
    fPart = UCase$(Left$(Trim$(StrValue), 2))
    IsUploadable = True
    If fPart = "GJ" Then
    ElseIf fPart = "AP" Then
    ElseIf fPart = "GJ" Then
    ElseIf fPart = "DT" Then
    ElseIf fPart = "CR" Then
    ElseIf fPart = "DR" Then
    ElseIf fPart = "R " Then
    ElseIf fPart = "O " Then
    ElseIf fPart = "P " Then
    ElseIf fPart = "I " Then
    ElseIf fPart = "TK" Then
    ElseIf fPart = "CV" Then
    ElseIf fPart = "PO" Then
    ElseIf fPart = "BR" Then
    ElseIf fPart = "BD" Then
    ElseIf fPart = "PS" Then
    ElseIf fPart = "TA" Then
    ElseIf fPart = "DB" Then
    ElseIf UCase$(Left$(StrValue, Len("MATCHING FIELD: CLOSING BALANCE"))) = "MATCHING FIELD: CLOSING BALANCE" Then
    ElseIf UCase$(Left$(StrValue, Len("MATCHING FIELD: OPENING BALANCE"))) = "MATCHING FIELD: OPENING BALANCE" Then
    Else
        IsUploadable = False
    End If
    Err.Clear
End Function
Public Function ExtractNumbers(ByVal StrValue As String) As String
    On Error Resume Next
    Dim i As Long
    Dim sResult As String
    Dim iLen As Long
    Dim myStr As String
    sResult = ""
    iLen = Len(StrValue)
    For i = 1 To iLen
        myStr = Mid$(StrValue, i, 1)
        If InStr("0123456789", myStr) > 0 Then
            sResult = sResult & myStr
        End If
    Next
    ExtractNumbers = sResult
    Err.Clear
End Function
Function StringPart(ByVal StrValue As String, Optional ByVal PartPosition As Long = 1, Optional ByVal Delimiter As String = ",", Optional TrimValue As Boolean = True) As String
    On Error Resume Next
    Dim xResult As String
    Dim xArray() As String
    If Len(StrValue) = 0 Then Exit Function
    xArray = Split(StrValue, Delimiter)
    Select Case PartPosition
    Case -1
        PartPosition = UBound(xArray) + 1
    Case 0
        PartPosition = 1
    End Select
    xResult = xArray(PartPosition - 1)
    If TrimValue = True Then
        xResult = Trim$(xResult)
    End If
    StringPart = xResult
    Err.Clear
End Function
Public Sub IsOnTop(ByVal wHandle As Long)
    On Error Resume Next
    Call SetWindowPos(wHandle, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Err.Clear
End Sub
Function File_RemoveBlankAndTrim(frmObj As Form, ByVal strFile As String) As String
    On Error Resume Next
    Dim lngFile As Long
    Dim rsStr As String
    Dim lngTmp As Long
    Dim tmpFile As String
    Dim rsTot As Long
    Dim rsCnt As Long
    rsTot = FileLen(strFile)
    StatusMessage frmObj, "Trimming and removing blank lines..."
    tmpFile = Left$(FileToken(strFile, "p"), 2) & "\tmp.txt"
    lngFile = FreeFile
    Open strFile For Input Access Read As #lngFile
        lngTmp = FreeFile
        Open tmpFile For Output Access Write As #lngTmp
            Do Until EOF(lngFile)
                Line Input #lngFile, rsStr
                rsStr = Replace$(rsStr, Quote, "")
                rsStr = Trim$(rsStr)
                If Len(rsStr) > 0 Then
                    Print #lngTmp, rsStr
                End If
            Loop
        Close lngTmp
    Close lngFile
    StatusMessage frmObj
    File_RemoveBlankAndTrim = tmpFile
    'Err.Clear
    Err.Clear
End Function
Public Sub LstViewRemoveDuplicates(LstView As ListView)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim spLines() As String
    Dim newCol As New Collection
    Dim spStr As String
    lstTot = LstView.ListItems.Count
    For lstCnt = 1 To lstTot
        spLines = LstViewGetRow(LstView, lstCnt)
        spStr = MvFromArray(spLines, RM)
        newCol.Add spStr, spStr
    Next
    LstView.ListItems.Clear
    lstTot = newCol.Count
    For lstCnt = 1 To lstTot
        Call StrParse(spLines, newCol.Item(lstCnt), RM)
        LstViewUpdate spLines, LstView, ""
    Next
    Set newCol = Nothing
    Err.Clear
End Sub
Public Function MvQuote(ByVal strData As String, Optional ByVal Delim As String = ",", Optional RemoveQuote As Boolean = True) As String
    On Error Resume Next
    Dim sData() As String
    Dim tCnt As Integer
    Dim wCnt As Integer
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(sData, strData, Delim)
    wCnt = UBound(sData)
    For tCnt = 1 To wCnt
        If RemoveQuote = True Then sData(tCnt) = Replace$(sData(tCnt), "'", "")
        sData(tCnt) = StringAdd("'", sData(tCnt), "'")
    Next
    MvQuote = MvFromArray(sData, Delim)
    Err.Clear
End Function
Public Sub LoadHeriachyPrefixed(ByVal SourceTable As String, SourceFld As String, Prefix As String, PrefixDelim As String, lstReport As ListView, Optional ByVal Icon As String = "close", Optional MyQuery As String = "")
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sFld As String
    Dim lPos As Long
    Dim tFlds As Long
    Dim cFlds As Long
    Dim vFlds As String
    With lstReport
        .ListItems.Clear
        .View = lvwList
        .GridLines = False
        .Checkboxes = False
        .Sorted = False
    End With
    If Len(MyQuery) = 0 Then
        Set tb = OpenRs("select distinct " & SQLQuote(SourceFld) & " from `" & SourceTable & "` order by " & SQLQuote(SourceFld))
    Else
        Set tb = OpenRs(MyQuery)
    End If
    tFlds = MvCount(SourceFld, ",")
    rsTot = AffectedRecords
    LstViewSetMemory lstReport, rsTot
    For rsCnt = 1 To rsTot
        vFlds = ""
        For cFlds = 1 To tFlds
            sFld = MvField(SourceFld, cFlds, ",")
            sFld = Trim$(MyRN(tb.Fields(sFld)))
            vFlds = vFlds & sFld & PrefixDelim
        Next
        vFlds = UCase$(RemDelim(vFlds, PrefixDelim))
        If Len(vFlds) = 0 Then vFlds = "<Blank>"
        lPos = LstViewFindItem(lstReport, vFlds, search_Text, search_Whole)
        If lPos = 0 Then
            LstViewAdd lstReport, vFlds, Prefix & "," & vFlds, "", 0, , Icon, Icon, True
        End If
        tb.MoveNext
    Next
    tb.Close
    Set tb = Nothing
    lstReport.Tag = Prefix
    Err.Clear
End Sub
Public Function RemDelim(ByVal Dataobj As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Dataobj = EnsureRight(Dataobj)
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delim)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delim
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function
Public Function LstViewAdd(lstReport As ListView, ByVal Text As String, Optional ByVal Key As String = "", Optional ByVal Tag As String = "", Optional ByVal Index As Long = 0, Optional ByVal Path As String = "", Optional ByVal Icon As String = "", Optional ByVal SmallIcon As String = "", Optional UpperCase As Boolean = False) As Variant
    On Error Resume Next
    Dim strX As String
    If Len(Icon) = 0 Then Icon = "closed"
    If Len(SmallIcon) = 0 Then SmallIcon = "opened"
    Text = ProperCase(Text)
    If UpperCase = True Then Text = UCase$(Text)
    Key = FixText(Key)
    Tag = FixText(Tag)
    Path = FixText(Path)
    strX = StringToMv(FM, Key, Tag, CStr(Index), Path)
    Set LstViewAdd = lstReport.ListItems.Add(, , Text, Icon, SmallIcon)
    If Len(Tag) > 0 Then LstViewAdd.Tag = strX
    If Len(Key) > 0 Then LstViewAdd.Key = Key
    Err.Clear
End Function
Public Function EnsureRight(ByVal Objdata As String) As String
    On Error Resume Next
    EnsureRight = Replace$(Objdata, Chr$(221), VM)
    EnsureRight = Replace$(EnsureRight, Chr$(222), Chr$(254))
    Err.Clear
End Function
Public Sub LoadHeriachy(ByVal SourceTable As String, SourceFld As String, Prefix As String, lstReport As ListView, Optional ByVal Icon As String = "close", Optional UpperCase As Boolean = True, Optional ConvertDistrict As Boolean = False, Optional MyQuery As String = "", Optional DataPrefix As String = "")
    On Error Resume Next
    Dim tb As New ADODB.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sFld As String
    Dim lPos As Long
    With lstReport
        .ListItems.Clear
        .View = lvwList
        .GridLines = False
        .Checkboxes = False
        .Sorted = False
    End With
    If Len(MyQuery) = 0 Then
        Set tb = OpenRs("select distinct " & SQLQuote(SourceFld) & " from `" & SourceTable & "` order by " & SQLQuote(SourceFld) & ";")
    Else
        Set tb = OpenRs(MyQuery)
    End If
    rsTot = AffectedRecords
    LstViewSetMemory lstReport, rsTot
    For rsCnt = 1 To rsTot
        sFld = Trim$(MyRN(tb.Fields(SourceFld)))
        If ConvertDistrict = True Then sFld = DistrictCode(sFld) & " - " & sFld
        If UpperCase = True Then sFld = UCase$(sFld)
        If Len(sFld) = 0 Then sFld = "<Blank>"
        If Len(DataPrefix) > 0 Then
            sFld = DataPrefix & sFld
        End If
        lPos = LstViewFindItem(lstReport, sFld, search_Text, search_Whole)
        If lPos = 0 Then
            LstViewAdd lstReport, sFld, Prefix & "," & sFld, "", 0, , Icon, Icon, UpperCase
        End If
        tb.MoveNext
    Next
    tb.Close
    Set tb = Nothing
    lstReport.Tag = Prefix
    Err.Clear
End Sub
Public Function StartEndDate(ByVal Yyyymm As String, Optional ByVal Sread As String = "") As String
    On Error Resume Next
    Dim smm As String
    Dim syy As String
    Dim sDate As String
    Dim eyy As String
    Dim emm As String
    Dim eDate As String
    Dim dLen As Integer
    If Len(Sread) = 0 Then
        Sread = "B"
    End If
    dLen = Len(Yyyymm) - 2
    smm = Right$(Yyyymm, 2)
    syy = Left$(Yyyymm, dLen)
    sDate = ToDate(StringToMv("/", "01", smm, syy))
    emm = Val(smm) + 1
    eyy = syy
    Select Case Val(emm)
    Case Is > 12
        emm = Val(emm) - 12
        eyy = Val(eyy) + 1
        If Val(eyy) = 100 Then
            eyy = "00"
        End If
    End Select
    eDate = ToDate(StringToMv("/", "01", emm, eyy))
    eDate = CDate(DateAdd("d", -1, CDate(eDate)))
    eDate = ToDate(eDate)
    Select Case UCase$(Left$(Sread, 1))
    Case "S"
        StartEndDate = sDate
    Case "E"
        StartEndDate = eDate
    Case "B"
        StartEndDate = StringAdd(sDate, ",", eDate)
    End Select
    Err.Clear
End Function
Public Function SwapDate(ByVal StrValue As String, Optional ConvertMySQL As Boolean = True) As String
    On Error Resume Next
    Dim SY As String
    Dim SM As String
    Dim SD As String
    StrValue = Format$(StrValue, "dd/mm/yyyy")
    SY = MvField(StrValue, 3, "/")
    SM = MvField(StrValue, 2, "/")
    SD = MvField(StrValue, 1, "/")
    StrValue = SM & "/" & SD & "/" & SY
    If ConvertMySQL = True Then
        SwapDate = SY & "-" & SM & "-" & SD
    Else
        SwapDate = StrValue
    End If
    Err.Clear
End Function
Public Sub LstViewFromMv(LstView As ListView, ByVal StringMv As String, Optional ByVal Delim As String = "", Optional ByVal lstClear As String = "")
    On Error Resume Next
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(spLine, StringMv, Delim)
    spTot = UBound(spLine)
    If Len(lstClear) = 0 Then
        LstView.ListItems.Clear
    End If
    For spCnt = 1 To spTot
        Call LstView.ListItems.Add(, , spLine(spCnt))
    Next
    LstViewAutoResize LstView
    Err.Clear
End Sub
Public Function TreeViewSearchPath(objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim iTmp As Long
    Dim iIndex As Long
    Dim mnIndex As Long
    Dim sCur As String
    TreeViewSearchPath = 0
    If objTree.Nodes.Count = 0 Then
    Err.Clear
        Exit Function
    End If
    mnIndex = 1
    'get the index of the root node that is at the top of the treeview
    iIndex = objTree.Nodes(mnIndex).FirstSibling.Index
    iTmp = iIndex
    sCur = FixText(objTree.Nodes(iIndex).FullPath)
    StrSearch = FixText(StrSearch)
    Select Case StrSearch
    Case sCur
        TreeViewSearchPath = objTree.Nodes(iIndex).Index
    Err.Clear
        Exit Function
    End Select
    If objTree.Nodes(iIndex).Children > 0 Then
        TreeViewSearchPath = TreeViewSearchChildPath(iIndex, objTree, StrSearch)
        If TreeViewSearchPath >= 1 Then
    Err.Clear
            Exit Function
        End If
    End If
    While iIndex <> objTree.Nodes(iTmp).LastSibling.Index
        'loop through all the root nodes
        sCur = FixText(objTree.Nodes(iIndex).Next.FullPath)
        Select Case StrSearch
        Case sCur
            TreeViewSearchPath = objTree.Nodes(iIndex).Next.Index
    Err.Clear
            Exit Function
        End Select
        If objTree.Nodes(iIndex).Next.Children > 0 Then
            TreeViewSearchPath = TreeViewSearchChildPath(objTree.Nodes(iIndex).Next.Index, objTree, StrSearch)
            If TreeViewSearchPath >= 1 Then
    Err.Clear
                Exit Function
            End If
        End If
        ' Move to the Next root Node
        iIndex = objTree.Nodes(iIndex).Next.Index
    Wend
    Err.Clear
End Function
Private Function TreeViewSearchChildPath(ByVal iNodeIndex As Long, ByVal objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim i As Long
    Dim iTempIndex As Long
    Dim lngChild As Long
    Dim sCur As String
    TreeViewSearchChildPath = 0
    StrSearch = FixText(StrSearch)
    iTempIndex = objTree.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    lngChild = objTree.Nodes(iNodeIndex).Children
    For i = 1 To lngChild
        sCur = FixText(objTree.Nodes(iTempIndex).FullPath)
        Select Case StrSearch
        Case sCur
            TreeViewSearchChildPath = objTree.Nodes(iTempIndex).Index
            Exit For
        End Select
        ' If the Node we are on has a child call the Sub again
        If objTree.Nodes(iTempIndex).Children > 0 Then
            TreeViewSearchChildPath = TreeViewSearchChildPath(iTempIndex, objTree, StrSearch)
            If TreeViewSearchChildPath >= 1 Then
                Exit For
            End If
        End If
        ' If we are not on the last child move to the next child Node
        If i <> objTree.Nodes(iNodeIndex).Children Then
            iTempIndex = objTree.Nodes(iTempIndex).Next.Index
        End If
    Next
    Err.Clear
End Function
Public Function TreeViewSearchText(objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim iTmp As Long
    Dim iIndex As Long
    Dim mnIndex As Long
    Dim sCur As String
    TreeViewSearchText = 0
    If objTree.Nodes.Count = 0 Then
    Err.Clear
        Exit Function
    End If
    mnIndex = 1
    'get the index of the root node that is at the top of the treeview
    iIndex = objTree.Nodes(mnIndex).FirstSibling.Index
    iTmp = iIndex
    sCur = FixText(objTree.Nodes(iIndex).Text)
    StrSearch = FixText(StrSearch)
    Select Case StrSearch
    Case sCur
        TreeViewSearchText = objTree.Nodes(iIndex).Index
    Err.Clear
        Exit Function
    End Select
    If objTree.Nodes(iIndex).Children > 0 Then
        TreeViewSearchText = TreeViewSearchChildText(iIndex, objTree, StrSearch)
        If TreeViewSearchText >= 1 Then
    Err.Clear
            Exit Function
        End If
    End If
    While iIndex <> objTree.Nodes(iTmp).LastSibling.Index
        'loop through all the root nodes
        sCur = FixText(objTree.Nodes(iIndex).Next.Text)
        Select Case StrSearch
        Case sCur
            TreeViewSearchText = objTree.Nodes(iIndex).Next.Index
    Err.Clear
            Exit Function
        End Select
        If objTree.Nodes(iIndex).Next.Children > 0 Then
            TreeViewSearchText = TreeViewSearchChildText(objTree.Nodes(iIndex).Next.Index, objTree, StrSearch)
            If TreeViewSearchText >= 1 Then
    Err.Clear
                Exit Function
            End If
        End If
        ' Move to the Next root Node
        iIndex = objTree.Nodes(iIndex).Next.Index
    Wend
    Err.Clear
End Function
Private Function TreeViewSearchChildText(ByVal iNodeIndex As Long, ByVal objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim i As Long
    Dim iTempIndex As Long
    Dim lngChild As Long
    Dim sCur As String
    TreeViewSearchChildText = 0
    StrSearch = FixText(StrSearch)
    iTempIndex = objTree.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    lngChild = objTree.Nodes(iNodeIndex).Children
    For i = 1 To lngChild
        sCur = FixText(objTree.Nodes(iTempIndex).Text)
        Select Case StrSearch
        Case sCur
            TreeViewSearchChildText = objTree.Nodes(iTempIndex).Index
            Exit For
        End Select
        ' If the Node we are on has a child call the Sub again
        If objTree.Nodes(iTempIndex).Children > 0 Then
            TreeViewSearchChildText = TreeViewSearchChildText(iTempIndex, objTree, StrSearch)
            If TreeViewSearchChildText >= 1 Then
                Exit For
            End If
        End If
        ' If we are not on the last child move to the next child Node
        If i <> objTree.Nodes(iNodeIndex).Children Then
            iTempIndex = objTree.Nodes(iTempIndex).Next.Index
        End If
    Next
    Err.Clear
End Function
Public Function MvField(ByVal strData As String, Optional ByVal fldPos As Long = 1, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim spData() As String
    Dim spCnt As Long
    MvField = ""
    If Len(Delim) = 0 Then Delim = VM
    If Len(strData) = 0 Then Exit Function
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function
Public Function FileToken(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    FileToken = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        FileToken = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        FileToken = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            FileToken = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Public Function RemoveDelim(ByVal strData As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    intDataSize = Len(strData)
    intDelimSize = Len(Delim)
    strLast = Right$(strData, intDelimSize)
    Select Case strLast
    Case Delim
        RemoveDelim = Left$(strData, (intDataSize - intDelimSize))
    Case Else
        RemoveDelim = strData
    End Select
    Err.Clear
End Function
Public Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = ";") As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = 0
    If Len(StringMv) = 0 Then
        MvSearch = 0
    Err.Clear
        Exit Function
    End If
    Call StrParse(TheFields, StringMv, Delim)
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Err.Clear
End Function
Public Function FileName_Validate(ByVal StrValue As String) As String
    On Error Resume Next
    Dim fPath As String
    Dim fFileN As String
    Dim fExt As String
    fPath = FileToken(StrValue, "p")
    fFileN = FileToken(StrValue, "fo")
    fExt = FileToken(StrValue, "e")
    fFileN = Replace$(fFileN, "\", "")
    fFileN = Replace$(fFileN, "/", "")
    fFileN = Replace$(fFileN, ":", "")
    fFileN = Replace$(fFileN, "*", "")
    fFileN = Replace$(fFileN, "?", "")
    fFileN = Replace$(fFileN, Quote, "")
    fFileN = Replace$(fFileN, "<", "")
    fFileN = Replace$(fFileN, ">", "")
    fFileN = Replace$(fFileN, "|", "")
    FileName_Validate = fPath & "\" & fFileN & "." & fExt
    Err.Clear
End Function
Public Function MvFromArray(vArray As Variant, Optional ByVal Delim As String = "", Optional StartingAt As Long = 1, Optional TrimItem As Boolean = True) As String
    On Error Resume Next
    If Len(Delim) = 0 Then Delim = VM
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    totArray = UBound(vArray)
    For i = StartingAt To totArray
        strL = vArray(i)
        If TrimItem = True Then strL = Trim$(strL)
        If i = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delim
        End If
    Next
    MvFromArray = BldStr
    Err.Clear
End Function
Public Function MvFromStartToEndOfArray(vArray As Variant, Optional StartingAt As Long = 1, Optional EndingAt As Long = -1, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    If Len(Delim) = 0 Then Delim = VM
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    totArray = UBound(vArray)
    If EndingAt <> -1 Then totArray = EndingAt
    For i = StartingAt To totArray
        strL = vArray(i)
        If i = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delim
        End If
    Next
    MvFromStartToEndOfArray = BldStr
    Err.Clear
End Function
Public Sub FileUpdate(ByVal filName As String, ByVal filLines As String, Optional ByVal Wora As String = "write")
    On Error Resume Next
    Dim iFileNum As Integer
    Dim cDir As String
    cDir = FileToken(filName, "p")
    MakeDirectory cDir
    iFileNum = FreeFile
    Select Case LCase$(Left$(Wora, 1))
    Case "w"
        Open filName For Output As #iFileNum
        Case "a"
            Open filName For Append As #iFileNum
            End Select
            Print #iFileNum, filLines
        Close #iFileNum
    Err.Clear
End Sub
Public Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    ArrayTot = UBound(varArray)
    StrSearch = LCase$(Trim$(StrSearch))
    ArraySearch = 0
    For arrayCnt = 1 To ArrayTot
        strCur = LCase$(varArray(arrayCnt))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
    Err.Clear
End Function
Public Function StrParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    If Len(Delim) = 0 Then Delim = VM
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
Public Function ProperCase(ByVal StrString As String, Optional Delim As String = "\") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spSubs() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim subTot As Long
    Dim subCnt As Long
    StrString = Trim$(StrString)
    spTot = StrParse(spItems, StrString, Delim)
    For spCnt = 1 To spTot
        spItems(spCnt) = StrConv(spItems(spCnt), vbProperCase)
    Next
    ProperCase = MvFromArray(spItems, Delim)
    Erase spItems
    Erase spSubs
    Err.Clear
End Function
Public Sub SetCounters()
    On Error Resume Next
    Dim lastCounter As Long
    lastCounter = MaxValue("Contacts", "id")
    WriteRecordMv "Param", "TableName", "Contacts", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("User Errors", "Transaction")
    WriteRecordMv "Param", "TableName", "User Errors", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Petty Cash", "id")
    WriteRecordMv "Param", "TableName", "Petty Cash", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Auditors", "id")
    WriteRecordMv "Param", "TableName", "Auditors", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Attachments", "id")
    WriteRecordMv "Param", "TableName", "Attachments", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Movement", "id")
    WriteRecordMv "Param", "TableName", "Movement", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Documents", "id")
    WriteRecordMv "Param", "TableName", "Documents", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Archived", "Transaction")
    WriteRecordMv "Param", "TableName", "Archived", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Recycled", "Transaction")
    WriteRecordMv "Param", "TableName", "Recycled", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Routes", "RouteID")
    WriteRecordMv "Param", "TableName", "Routes", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Drop Off Points", "doID")
    WriteRecordMv "Param", "TableName", "Drop Off Points", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Pick Up Points", "puID")
    WriteRecordMv "Param", "TableName", "Pick Up Points", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Service Providers", "spID")
    WriteRecordMv "Param", "TableName", "Service Providers", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Specimen Signatures", "ID")
    WriteRecordMv "Param", "TableName", "Specimen Signatures", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Claims", "ClaimID")
    WriteRecordMv "Param", "TableName", "Claims", "Sequence", CStr(lastCounter)
    lastCounter = MaxValue("Cellphone Simswaprequests", "ID")
    WriteRecordMv "Param", "TableName", "Cellphone Simswaprequests", "Sequence", CStr(lastCounter)
    Err.Clear
End Sub
Private Function TotalUpdates() As Long
    On Error Resume Next
    Dim myVer As String
    Dim serverPth As String
    Dim serverExe As String
    Dim iU As Integer
    DependencyCount
    SaveReg "server", "z:\eFas\LiveUpdate\eFas.exe", "account", "eFas"
    ' check the server path
    myVer = FileVersion(App.Path & "\eFas.exe")
    serverPth = ReadReg("server", "account", "eFas")
    iU = 0
    If Len(serverPth) > 0 Then
        If FileExists(serverPth) = True Then
            serverExe = FileVersion(serverPth)
            serverExe = Replace$(serverExe, ".", "")
            myVer = Replace$(myVer, ".", "")
            If Val(serverExe) > Val(myVer) Then
                iU = iU + 1
            End If
        End If
    End If
    Dim tDep As Long
    Dim cDep As Long
    Dim sDep As String
    Dim nDep As String
    ' how many dependancies do we have to update
    tDep = Val(ReadReg("dependencies", "account", "eFas"))
    For cDep = 1 To tDep
        sDep = ReadReg("dependency" & cDep, "account", "eFas")
        If Len(sDep) > 0 Then
            If FileExists(sDep) = True Then
                serverExe = FileVersion(sDep)
                nDep = GetSysDir & "\" & FileToken(sDep, "f")
                If FileExists(nDep) = True Then
                    myVer = FileVersion(nDep)
                    serverExe = Replace$(serverExe, ".", "")
                    myVer = Replace$(myVer, ".", "")
                    If Val(serverExe) > Val(myVer) Then
                        iU = iU + 1
                    End If
                End If
            End If
        End If
    Next
    TotalUpdates = iU
    Err.Clear
End Function
Public Sub ImportExistingPayments(frmObj As Form, ByVal strFile As String)
    On Error Resume Next
    Dim db As DAO.Database
    Dim rsT As ADODB.Recordset
    Dim Rs As DAO.Recordset
    Dim sheetNames As String
    Dim spTot As Long
    Dim spSheetNames() As String
    Dim spCnt As Long
    Dim rsColumns As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim spName As String
    sheetNames = ExcelSheetNames(strFile, VM)
    Set db = DAO.OpenDatabase(strFile, False, True, "Excel 8.0; HDR=NO;")
    spTot = StrParse(spSheetNames, sheetNames, VM)
    For spCnt = 1 To spTot
        Set Rs = db.OpenRecordset(spSheetNames(spCnt) & "$")
        rsColumns = DaoFldNames(Rs)
        If MvCount(rsColumns, ",") < 1 Then GoTo NextSheet
        rsTot = Rs.RecordCount
        ProgBarInit frmObj.progBar, rsTot
        StatusMessage frmObj, "Updating payment status from " & spSheetNames(spCnt)
        For rsCnt = 1 To rsTot
            frmObj.progBar.Value = rsCnt
            spName = UCase$(Trim$(RN(Rs(0))))
            If Len(spName) > 0 Then
                Execute "update `t&s advance dom ca audittrail` set HasPayment = 'Y' where funcarea = 'AP" & spName & "';"
            End If
            DoEvents
            Rs.MoveNext
        Next
        StatusMessage frmObj
        ProgBarClose frmObj.progBar
NextSheet:
    Next
    db.Close
    Err.Clear
End Sub
Public Function LastThreeDigits(frmObj As Form, ByVal sTable As String) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim Scode As String
    Dim rsItems As ADODB.Recordset
    Dim rsCol As Collection
    Set rsCol = New Collection
    Set rsItems = OpenRs("select * from `" & sTable & "` where PostingLevel = 'Y'")
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Checking " & sTable & " code structure..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        Scode = MyRN(rsItems.Fields("Code"))
        Scode = Right$(Scode, 3)
        rsCol.Add Scode, Scode
        DoEvents
    Next
    LastThreeDigits = MvFromCollection(rsCol, ",")
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Err.Clear
End Function
Public Function ObjectiveProgramme(ByVal StrObjective As String) As String
    On Error Resume Next
    Dim openPos As Long
    Dim strRest As String
    openPos = InStrRev(StrObjective, "(", , vbTextCompare)
    If openPos > 0 Then
        strRest = Mid$(StrObjective, openPos + 1)
        strRest = Replace$(strRest, ")", "")
        ObjectiveProgramme = Trim$(ExtractNumbers(strRest))
    Else
        ObjectiveProgramme = "0"
    End If
    Err.Clear
End Function
Public Sub DaoCreateIndexes(ByVal DbName As String, ByVal TbName As String, ByVal IndexNames As String)
    On Error Resume Next
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idxLoop As DAO.Index
    Dim IdxName As String
    Dim spIndexes() As String
    Dim spTot As Integer
    Dim spCnt As Integer
    Call StrParse(spIndexes, IndexNames, ",")
    spTot = UBound(spIndexes)
    Set dbs = DAO.OpenDatabase(DbName)
    Set tdf = dbs.TableDefs(TbName)
    With tdf
        For spCnt = 1 To spTot
            IdxName = spIndexes(spCnt)
            Set idxLoop = .CreateIndex(IdxName)
            idxLoop.Fields.Append idxLoop.CreateField(IdxName)
            .Indexes.Append idxLoop
        Next
        .Indexes.Refresh
    End With
    dbs.Close
    Set dbs = Nothing
    Err.Clear
End Sub
Public Sub TreeViewCheckChildNodes(oParentNode As Node, ByVal bChecked As Boolean)
    On Error Resume Next
    Dim oNode As Node
    ' Get the first child node
    Set oNode = oParentNode.Child
    ' Loop through the child nodes of this node
    ' until there are none left...
    Do While Not oNode Is Nothing
        ' Check/Uncheck the node
        oNode.Checked = bChecked
        'Call this function again for the
        ' child node, so that it's child nodes
        ' can get checked/unchecked.
        TreeViewCheckChildNodes oNode, bChecked
        ' Get the next child node of this node
        Set oNode = oNode.Next
    Loop
    Err.Clear
End Sub
Public Function TreeViewKeys(objTree As TreeView, CheckedStatus As Boolean, Optional Delim As String = ",") As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsKey As String
    Dim rsEnd As String
    rsEnd = ""
    rsTot = objTree.Nodes.Count
    For rsCnt = 1 To rsTot
        If objTree.Nodes(rsCnt).Checked = CheckedStatus Then
            rsKey = objTree.Nodes(rsCnt).Key
            rsEnd = rsEnd & rsKey & Delim
        End If
    Next
    TreeViewKeys = RemoveDelim(rsEnd, Delim)
    Err.Clear
End Function
Public Function IsThereYN(StrContainer As String, StrSearchFor As String, Optional Delim As String = ",") As String
    On Error Resume Next
    If MvSearch(StrContainer, StrSearchFor, Delim) = 0 Then
        IsThereYN = "Y"
    Else
        IsThereYN = "N"
    End If
    Err.Clear
End Function
Public Sub TreeViewCheckKeys(objTree As MSComctlLib.TreeView, theModules As String)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsKey As String
    Dim rsEnd As String
    rsTot = objTree.Nodes.Count
    For rsCnt = 1 To rsTot
        rsKey = objTree.Nodes(rsCnt).Key
        If MvSearch(theModules, rsKey, ",") > 0 Then
            objTree.Nodes(rsCnt).Checked = False
        Else
            objTree.Nodes(rsCnt).Checked = True
        End If
    Next
    objTree.Refresh
    Err.Clear
End Sub
Function MvSort_String(vString As String, Optional Delim As String = "") As String
    On Error Resume Next
    Dim lLoop1 As Long
    Dim lHold As Long
    Dim lHValue As String
    Dim lTemp As String
    Dim vArray As Variant
    vArray = Split(vString, Delim)
    lHValue = LBound(vArray)
    Do
        lHValue = 3 * lHValue + 1
    Loop Until lHValue > UBound(vArray)
    Do
        lHValue = lHValue / 3
        For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
            lTemp = vArray(lLoop1)
            lHold = Val(lLoop1)
            Do While vArray(lHold - lHValue) > lTemp
                vArray(lHold) = vArray(lHold - lHValue)
                lHold = lHold - lHValue
                If lHold < lHValue Then Exit Do
            Loop
            vArray(lHold) = lTemp
        Next
    Loop Until lHValue = LBound(vArray)
    MvSort_String = MvFromArray(vArray, Delim, 0)
    Err.Clear
End Function
Function MvSort(vString As String, Optional Delim As String = "") As String
    On Error Resume Next
    MvSort = MvSort_String(vString, Delim)
    Err.Clear
End Function
Public Function isValidEmail(ByVal myEmail As String) As Boolean
    On Error GoTo myError
    isValidEmail = True
    If Len(myEmail) = 0 Then Exit Function
    isValidEmail = ValidEmail(myEmail)
    If isValidEmail = False Then Exit Function
    ' just as a second check
    Dim myRegExp As RegExp
    Dim myMatches As MatchCollection
    Set myRegExp = New RegExp ' Create Regular expresion To extract valid email addresses
    myRegExp.Pattern = "[a-zA-Z0-9-_.]+@[a-zA-Z0-9-_.]+\.[a-zA-Z0-9]+" ' Set pattern.
    myRegExp.IgnoreCase = False ' Set case insensitivity.
    myRegExp.Global = True ' Set global applicability.
    'Execute search.
    Set myMatches = myRegExp.Execute(myEmail)
    If myMatches.Count <> 1 Then
        isValidEmail = False 'i want To make sure is Set to false ONLY If found a bad address
    End If
    Set myMatches = Nothing
    Err.Clear
    Exit Function
myError:
    Err.Clear
End Function
Function ValidEmail(ByVal strCheck As String) As Boolean
    On Error Resume Next
    Dim bCK As Boolean
    Dim strDomainType As String
    '    Dim strDomainName As String
    Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
    Dim i As Integer
    bCK = Not InStr(1, strCheck, Quote) > 0 'Check to see if there is a double Quote
    If Not bCK Then GoTo ExitFunction
    bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
    If Not bCK Then GoTo ExitFunction
    ' Check for invalid characters.
    If Len(strCheck) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(strCheck, Mid$(sInvalidChars, i, 1)) > 0 Then
                bCK = False
                GoTo ExitFunction
            End If
        Next
    Else
        For i = 1 To Len(strCheck)
            If InStr(sInvalidChars, Mid$(strCheck, i, 1)) > 0 Then
                bCK = False
                GoTo ExitFunction
            End If
        Next
    End If
    If InStr(1, strCheck, "@") > 1 Then
        'Check for an @ symbol
        bCK = Len(Left$(strCheck, InStr(1, strCheck, "@") - 1)) > 0
    Else
        bCK = False
    End If
    If Not bCK Then GoTo ExitFunction
    strCheck = Right$(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
    bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
    If Not bCK Then GoTo ExitFunction
    strDomainType = Right$(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
    bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
    If Not bCK Then GoTo ExitFunction
    bCK = InStr(1, strCheck, ".") > 0 'Check to see if there are 1 dot
    If Not bCK Then GoTo ExitFunction
    strCheck = Left$(strCheck, Len(strCheck) - Len(strDomainType) - 1)
    Do Until InStr(1, strCheck, ".") <= 1
        If Len(strCheck) >= InStr(1, strCheck, ".") Then
            strCheck = Left$(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
        Else
            bCK = False
            GoTo ExitFunction
        End If
    Loop
    If strCheck = "." Or Len(strCheck) = 0 Then bCK = False
ExitFunction:
    ValidEmail = bCK
    Err.Clear
End Function
Sub Outlook_CreateTask(ByVal StrSubject As String, ByVal StrOwner As String, ByVal StrBody As String, ByVal StrStartDate As String, ByVal strDueDate As String, ByVal SetReminder As Integer, ByVal StrReminderTime As String, ByVal StrReminderDate As String, ByVal StrRecepient As String, Optional ByVal StrPriority As String = "Normal", Optional ByVal StrStatus As String = "Not Completed", Optional ByVal strSensitivity As String = "Normal", Optional ByVal StrProfileEmail As String = "")
    On Error Resume Next
    Dim oApp As Outlook.Application
    Dim oNspc As Outlook.NameSpace
    Dim myTasks As Outlook.mapiFolder
    Dim myTask As Outlook.TaskItem
    Dim myRecipient As Outlook.Recipient
    Dim myRecipient1 As Outlook.Recipient
    Dim strFilter As String
    '    Dim outRun As Boolean
    'outRun = IIf(InStr(1, RunningProcesses(vm), "outlook.exe", vbTextCompare) > 0, True, False)
    'If outRun = True Then
    '    Set oApp = GetObject("", "Outlook.Application")
    'Else
    Set oApp = New Outlook.Application
    'End If
    Set oNspc = oApp.GetNamespace("MAPI")
    Select Case Len(StrProfileEmail)
    Case 0
        Set myTasks = oNspc.GetDefaultFolder(Outlook.olFolderTasks)
    Case Else
        Set myRecipient = oNspc.CreateRecipient(StrProfileEmail)
        myRecipient.Resolve
        If myRecipient.Resolved = True Then
            Set myTasks = oNspc.GetSharedDefaultFolder(myRecipient, Outlook.olFolderTasks)
        End If
    End Select
    If LCase$(TypeName(myTasks)) <> "mapifolder" Then
        RetAnswer = MyPrompt("An error was encountered trying to create the task.", "o", "e", "Outlook Create Task")
        GoTo FinishUp
    End If
    ' define the filters using subject
    strFilter = "[Subject] = '" & StrSubject & "'"
    ' search for the task
    Set myTask = myTasks.items.Find(strFilter)
    ' if appointment is not found, then add it to folder
    If TypeName(myTask) = "Nothing" Then
        Set myTask = myTasks.items.Add()
    End If
    GoSub LinkData
FinishUp:
    'If outRun = False Then oApp.Quit
    oApp.Quit
    DoEvents
    Set oApp = Nothing
    Set oNspc = Nothing
    Set myTasks = Nothing
    Set myTask = Nothing
    Set myRecipient = Nothing
    Set myRecipient1 = Nothing
    Err.Clear
    Exit Sub
LinkData:
    With myTask
        .Subject = StrSubject
        .Owner = StrOwner
        .Body = StrBody
        .dueDate = strDueDate
        .StartDate = StrStartDate
        .Status = StatusValue(StrStatus)
        .Importance = PriorityValue(StrPriority)
        .Sensitivity = SensitivityValue(strSensitivity)
        If SetReminder = 1 Then
            .ReminderSet = True
            .ReminderTime = StrReminderDate & " " & StrReminderTime
        Else
            .ReminderSet = False
        End If
        ' find if the recipient exists, if not add
        .Assign
        If RecipientSearch(myTask, StrRecepient) = -1 Then
            .Recipients.Add StrRecepient
        End If
        .Save
        If .Recipients.Count > 0 Then
            .Send
        End If
    End With
    Err.Clear
End Sub
Function StatusValue(ByVal strCurrent As String) As Integer
    On Error Resume Next
    strCurrent = LCase$(strCurrent)
    Select Case strCurrent
    Case "not started": StatusValue = 0
    Case "in progress": StatusValue = 1
    Case "completed": StatusValue = 2
    Case "waiting on someone": StatusValue = 3
    Case "deferred": StatusValue = 4
    End Select
    Err.Clear
End Function
Function PriorityValue(ByVal valCur As String) As Integer
    On Error Resume Next
    Select Case LCase$(valCur)
    Case "low": PriorityValue = 0
    Case "normal": PriorityValue = 1
    Case "high": PriorityValue = 2
    End Select
    Err.Clear
End Function
Function SensitivityValue(ByVal valCur As String) As Integer
    On Error Resume Next
    Select Case LCase$(valCur)
    Case "normal": SensitivityValue = 0
    Case "personal": SensitivityValue = 1
    Case "private": SensitivityValue = 2
    Case "confidential": SensitivityValue = 3
    End Select
    Err.Clear
End Function
Public Sub LstBoxRemoveItemAPI(lstBox As Variant, ParamArray cboItems())
    On Error Resume Next
    Dim cboItem As Variant
    Dim cboPos As Long
    Dim cboStr As String
    Select Case TypeName(lstBox)
    Case "ListBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            cboPos = LstBoxFindExactItemAPI(lstBox, cboStr$)
            If cboPos <> -1 Then
                Call SendMessage(lstBox.hWnd, LB_DELETESTRING, cboPos, ByVal 0&)
            End If
        Next
    Case "ComboBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            cboPos = LstBoxFindExactItemAPI(lstBox, cboStr$)
            If cboPos <> -1 Then
                Call SendMessage(lstBox.hWnd, CB_DELETESTRING, cboPos, ByVal 0&)
            End If
        Next
    End Select
    Err.Clear
End Sub
Public Function FIS_FixAmount(ByVal StrValue As String) As String
    On Error Resume Next
    Dim sH As String
    Dim sC As String
    StrValue = Trim$(StrValue)
    If InStr(1, StrValue, ".") = 0 Then
        If StrValue = "-" Then StrValue = "0"
        FIS_FixAmount = StrValue & ".00"
    Else
        sH = MvField(StrValue, 1, ".")
        sC = MvField(StrValue, 2, ".")
        If Len(sC) = 1 Then
            sC = sC & "0"
        ElseIf Len(sC) = 0 Then
            sC = "00"
        End If
        FIS_FixAmount = sH & "." & sC
    End If
    Err.Clear
End Function
Public Function MvWriteValueAt(ByVal StringMv As String, ByVal intPos As Long, ByVal Strtowrite As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim TheFields() As String
    Dim MaxPos As Long
    MvWriteValueAt = StringMv
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(TheFields, StringMv, Delim)
    MaxPos = UBound(TheFields)
    Select Case intPos
    Case Is < 0
        MaxPos = MaxPos + 1
        ReDim Preserve TheFields(MaxPos)
        TheFields(MaxPos) = Strtowrite
    Case Is > MaxPos
        ReDim Preserve TheFields(intPos)
        TheFields(intPos) = Strtowrite
    Case Else
        TheFields(intPos) = Strtowrite
    End Select
    MvWriteValueAt = MvFromArray(TheFields, Delim)
    Err.Clear
End Function
Function ExcelFileToTextFile(ByVal strFile As String, Optional NewFileFormat As ExcelFileFormat = TextMSDOS) As String
    On Error Resume Next
    Err.Clear
    Exit Function
    Dim exlApp As Excel.Application
    Dim XLWkb As Excel.Workbook
    Dim newFile As String
    Set exlApp = New Excel.Application
    If TypeName(exlApp) = "Nothing" Then
        ExcelFileToTextFile = ""
    Err.Clear
        Exit Function
    End If
    exlApp.DisplayAlerts = False
    exlApp.ScreenUpdating = False
    exlApp.Workbooks.Open strFile
    exlApp.Visible = False
    exlApp.WindowState = xlMinimized
    Set XLWkb = exlApp.ActiveWorkbook
    Select Case NewFileFormat
    Case CSV
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".csv"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlCSV, CreateBackup:=False
    Case DBF4
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".dbf"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlDBF4, CreateBackup:=False
    Case Html
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".html"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlHtml, CreateBackup:=False
    Case TextMSDOS
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".txt"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlTextMSDOS, CreateBackup:=False
    Case TextPrinter
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".prn"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlTextPrinter, CreateBackup:=False
    Case TextWindows
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".txt"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlTextWindows, CreateBackup:=False
    Case XMLSpreadsheet = 11
        newFile = FileToken(strFile, "p") & "\" & FileToken(strFile, "fo") & ".xml"
        XLWkb.SaveAs Filename:=newFile, FileFormat:=xlXMLSpreadsheet, CreateBackup:=False
    End Select
    DoEvents
    exlApp.Quit
    Set exlApp = Nothing
    Set XLWkb = Nothing
    ExcelFileToTextFile = newFile
    Err.Clear
End Function
Public Sub MonthlyScheduleByFieldConsolidate(frmObj As Form, ByVal SourceTb As String, ByVal DateFld As String, ByVal amtFld As String, ByVal fldNames As String, ByVal FldNamesShown As String, LstView As ListView, ByVal sDate As String, ByVal eDate As String, Optional ByVal strSQL As String = "", Optional RemoveMonthsColumns As Boolean = False, Optional ByVal sqlAfter As String = "", Optional SQLAfterHeadings As String = "", Optional StartAlignment As Long = 2, Optional ByVal OrderBy As String = "")
    On Error Resume Next
    StatusMessage frmObj, "Preparing schedule, please be patient..."
    Dim tbF As ADODB.Recordset
    Dim tbS As ADODB.Recordset
    Dim nCnt As Long
    Dim sMonths As String
    Dim lsDate As Long
    Dim leDate As Long
    Dim fName As String
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim invoiceDate As String
    Dim xTot() As String
    Dim xFld() As String
    Dim fTot As Long
    Dim fCnt As Long
    Dim sql As String
    Dim spFldNames() As String
    Dim spShwNames() As String
    Dim totFlds As Long
    Dim cntFlds As Long
    Dim strName As String
    Dim StrValue As String
    Dim StrKey As String
    Dim strAmt As String
    Dim oldM As String
    Dim amtPos As Long
    Dim nAmt As String
    Dim tAmt As String
    Dim mMonths() As String
    lsDate = Val(DateIconv(sDate))
    leDate = Val(DateIconv(eDate))
    For nCnt = lsDate To leDate
        sMonths = DateOconv(CStr(nCnt))
        sMonths = Format$(sMonths, "mmm yyyy")
        ArrayUpdate mMonths, sMonths
    Next
    sMonths = MvFromArray(mMonths, ",")
    oldM = sMonths
    fTot = MvCount(sMonths, ",")
    ReDim xTot(fTot) As String
    Call StrParse(xFld, sMonths, ",")
    CreateTableWithIndexNames "MonthlySchedule", "Key," & fldNames & "," & sMonths & ",Amount", , , "Key", , , , , "Key"
    sql = "select * from `" & SourceTb & "` where `" & DateFld & "` >= '" & SwapDate(sDate) & "' and `" & DateFld & "` <= '" & SwapDate(eDate) & "'"
    If Len(strSQL) > 0 Then
        sql = strSQL
    End If
    Execute "delete from MonthlySchedule;"
    Set tbF = OpenRs(sql)
    rsTot = AffectedRecords
    ProgBarInit frmObj.progBar, rsTot
    Call StrParse(spFldNames, fldNames, ",")
    totFlds = UBound(spFldNames)
    Call StrParse(spShwNames, FldNamesShown, ",")
    StatusMessage frmObj, "Compiling monthly schedule..."
    For rsCnt = 1 To rsTot
        frmObj.progBar.Value = rsCnt
        invoiceDate = MyRN(tbF.Fields(DateFld))
        strAmt = MyRN(tbF.Fields(amtFld))
        strAmt = ProperCase(strAmt)
        fName = Format$(invoiceDate, "mmm yyyy")
        ' compile the key
        StrKey = ""
        For cntFlds = 1 To totFlds
            ' what is the field name
            strName = spFldNames(cntFlds)
            ' what is the value in the field
            StrValue = MyRN(tbF.Fields(strName))
            StrKey = StrKey & StrValue & RM
        Next
        StrKey = RemoveDelim(StrKey, RM)
        Set tbS = SeekRs("Key", StrKey, "MonthlySchedule")
        Select Case tbS.EOF
        Case True
            tbS.AddNew
            tbS.Fields(fName) = strAmt
            tbS.Fields("Amount") = strAmt
            tbS.Fields("Key") = StrKey
            For cntFlds = 1 To totFlds
                strName = spFldNames(cntFlds)
                StrValue = MyRN(tbF.Fields(strName))
                tbS.Fields(strName) = StrValue
            Next
            UpdateRs tbS
        Case Else
            nAmt = Val(MyRN(tbS.Fields(fName))) + Val(strAmt)
            nAmt = ProperAmount(nAmt)
            tAmt = Val(MyRN(tbS.Fields("Amount"))) + Val(strAmt)
            tAmt = ProperAmount(tAmt)
            tbS.Fields(fName) = nAmt
            tbS.Fields("Amount") = tAmt
            UpdateRs tbS
        End Select
        DoEvents
        tbF.MoveNext
    Next
    RemoveEmptyColumns "MonthlySchedule"
    ' delete the key field
    DeleteFields "MonthlySchedule", "Key"
    ' delete the months field
    If RemoveMonthsColumns = True Then
        fTot = MvCount(oldM, ",")
        For fCnt = 1 To fTot
            DeleteFields "MonthlySchedule", MvField(oldM, fCnt, ",")
        Next
    End If
    oldM = FieldNames("MonthlySchedule")
    oldM = Replace$(oldM, "Key", "")
    oldM = MvRemoveBlanks(oldM, ",")
    If Len(sqlAfter) = 0 Then
        StrKey = "select " & SQLQuote(oldM) & " from monthlyschedule"
        If Len(OrderBy) > 0 Then StrKey = StrKey & " order by " & OrderBy
        LstView.Sorted = False
        ViewSQLNew StrKey, LstView, fldNames, , , , , , , , , "Amount"
        LstViewSumColumns LstView, True, "Amount"
        Call LstViewAutoResize(LstView)
        LstView.Refresh
        If Len(FldNamesShown) > 0 Then
            rsTot = MvCount(FldNamesShown, ",")
            For rsCnt = 1 To rsTot
                fName = StringPart(FldNamesShown, rsCnt, ",")
                LstView.ColumnHeaders(rsCnt).Text = ProperCase(fName)
            Next
        End If
        For rsCnt = StartAlignment To LstView.ColumnHeaders.Count
            LstView.ColumnHeaders(rsCnt).Alignment = lvwColumnRight
        Next
    Else
        LstView.Sorted = False
        ViewSQLNew sqlAfter, LstView, SQLAfterHeadings, , , , , , , , , "Amount"
        LstViewSumColumns LstView, True, "Amount"
        Call LstViewAutoResize(LstView)
        LstView.Refresh
        If Len(FldNamesShown) > 0 Then
            rsTot = MvCount(FldNamesShown, ",")
            For rsCnt = 1 To rsTot
                fName = StringPart(FldNamesShown, rsCnt, ",")
                LstView.ColumnHeaders(rsCnt).Text = ProperCase(fName)
            Next
        End If
        For rsCnt = StartAlignment To LstView.ColumnHeaders.Count
            LstView.ColumnHeaders(rsCnt).Alignment = lvwColumnRight
        Next
    End If
    oldM = LstViewColNames(LstView)
    rsTot = UBound(mMonths)
    For rsCnt = 1 To rsTot
        strName = mMonths(rsCnt)
        amtPos = MvSearch(oldM, strName, ",")
        If amtPos > 0 Then
            LstViewSumColumns LstView, True, strName
        End If
    Next
    LstViewAutoResize LstView
    LstView.Refresh
    Err.Clear
End Sub
Public Sub ArrayUpdate(vArray() As String, ParamArray items())
    On Error Resume Next
    Dim Item As Variant
    Dim NewItem As String
    Dim aPos As Long
    Dim vTot As Long
    vTot = UBound(vArray)
    For Each Item In items
        NewItem = CStr(Item)
        aPos = ArraySearch(vArray, NewItem)
        If aPos = 0 Then
            vTot = vTot + 1
            ReDim Preserve vArray(vTot)
            vArray(vTot) = NewItem
        End If
    Next
    Err.Clear
End Sub
Public Sub Outlook_ImportGlobalAddressList(frmObj As Form)
    On Error Resume Next
    Dim cdoSession As MAPI.Session
    Dim rsEmail As ADODB.Recordset
    Dim olAL As MAPI.AddressList
    Dim olAE As MAPI.AddressEntry
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim strEntryID  As String
    Dim strEmail As String
    Dim strName As String
    Set cdoSession = New MAPI.Session
    cdoSession.Logon
    ' get the global address list
    Set olAL = cdoSession.GetAddressList(MAPI.CdoAddressListGAL)
    If TypeName(olAL) = "Nothing" Then
        Call MyPrompt("The global address list could not be accessed, please try again later.", "o", "e", "Global Address List Error")
    Err.Clear
        Exit Sub
    End If
    rsTot = olAL.AddressEntries.Count
    ProgBarInit frmObj.progBar, rsTot
    StatusMessage frmObj, "Importing global address list, please wait..."
    rsCnt = 0
    Execute "delete from `GlobalAddressList`;"
    For Each olAE In olAL.AddressEntries
        rsCnt = rsCnt + 1
        frmObj.progBar.Value = rsCnt
        strEntryID = olAE.ID
        strEmail = cdoSession.GetAddressEntry(strEntryID).Fields(&H39FE001E)
        strName = olAE.Name
        Set rsEmail = SeekRs("FullName", strName, "GlobalAddressList")
        If rsEmail.EOF = True Then rsEmail.AddNew
        rsEmail.Fields("FullName") = strName
        rsEmail.Fields("Email") = strEmail
        UpdateRs rsEmail
        DoEvents
    Next
    cdoSession.Logoff
    Set rsEmail = Nothing
    Set cdoSession = Nothing
    Set olAL = Nothing
    Set olAE = Nothing
    ProgBarClose frmObj.progBar
    StatusMessage frmObj
    Err.Clear
End Sub
Function ReadableFigure(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = Replace$(StrValue, ",", "")
    ReadableFigure = Replace$(StrValue, ".", "")
    Err.Clear
End Function
Sub SaveToWord(ByVal SavePath As String, ByVal StrCaption As String, lstReport As ListView, Optional FootNote As String = "", Optional ByVal xOrientation As String = "landscape")
    On Error Resume Next
    Dim xFile As String
    Dim xPath As String
    If lstReport.ListItems.Count = 0 Then Exit Sub
    xFile = SavePath & "\" & StrCaption & ".doc"
    xFile = FileName_Validate(xFile)
    xPath = FileToken(xFile, "p")
    If DirExists(xPath) = False Then MakeDirectory xPath
    LstViewToWordTable StrCaption, FootNote, xFile, lstReport, xOrientation
    DoEvents
    Err.Clear
End Sub
Public Function TreeViewCheckNormal(trView As TreeView, Optional CheckStatus As Boolean = False)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = trView.Nodes.Count
    For lstCnt = 1 To lstTot
        trView.Nodes(lstCnt).Checked = CheckStatus
    Next
    Err.Clear
End Function
Public Sub DaoTableNamesToLstView(Dbase As String, LstView As ListView, Optional ByVal boolClear As Boolean = True)
    On Error Resume Next
    'loads table names into a listview
    Dim db As DAO.Database
    Dim StrDt As String
    Dim rsCnt As Long
    Dim rsTot As Long
    ' clear the listview if specified
    If boolClear = True Then LstView.ListItems.Clear
    ' open the database
    Set db = DAO.OpenDatabase(Dbase)
    ' how many tables are there
    rsTot = db.TableDefs.Count - 1
    For rsCnt = 0 To rsTot
        DoEvents
        StrDt = db.TableDefs(rsCnt).Name
        Select Case StrDt
        Case "MSysAccessStorage", "MSysAccessXML", "MSysACEs", "MSysIMEXColumns", "MSysIMEXSpecs", "MSysObjects", "MSysQueries", "MSysRelationships", "MSysAccessObjects"
        Case Else
            LstView.ListItems.Add , , StrDt
        End Select
    Next
    db.Close
    Set db = Nothing
    Err.Clear
End Sub
Public Function MvRecord(tb As DAO.Recordset, totFlds As Long, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim cntFlds As Long
    Dim rsValues As String
    rsValues = ""
    For cntFlds = 0 To totFlds
        Select Case cntFlds
        Case totFlds
            rsValues = rsValues & MyRN(tb.Fields(cntFlds).Value)
        Case Else
            rsValues = rsValues & MyRN(tb.Fields(cntFlds).Value) & Delim
        End Select
    Next
    MvRecord = rsValues
    Err.Clear
End Function
Public Function PersalAmount(ByVal StrValue As String) As String
    On Error Resume Next
    Dim sCents As String
    Dim sRands As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsVal As String
    Dim mLeft As String
    Dim mRight As String
    rsStr = ""
    rsTot = Len(StrValue)
    For rsCnt = 1 To rsTot
        rsVal = Mid$(StrValue, rsCnt, 1)
        If InStr(1, "-0123456789", rsVal) > 0 Then
            rsStr = rsStr & rsVal
        End If
    Next
    rsStr = Trim$(rsStr)
    If Len(rsStr) = 0 Then rsStr = "0.00"
    rsTot = Len(rsStr)
    sCents = Right$(rsStr, 2)
    sRands = Left$(rsStr, rsTot - 2)
    PersalAmount = ProperAmount(sRands & "." & sCents)
    Err.Clear
End Function
Public Function PersalDate(ByVal StrValue As String) As String
    On Error Resume Next
    Dim sdd As String
    Dim smm As String
    Dim sYYYY As String
    Dim lenValue As Long
    lenValue = Len(StrValue)
    If lenValue = 0 Then StrValue = "00000000"
    lenValue = Len(StrValue)
    sdd = Right$(StrValue, 2)
    smm = Right$(StrValue, 4)
    smm = Left$(smm, 2)
    sYYYY = Left$(StrValue, lenValue - 4)
    sdd = sdd & "/" & smm & "/" & sYYYY
    If IsDate(sdd) = True Then
        PersalDate = sdd
    Else
        PersalDate = ""
    End If
    Err.Clear
End Function
