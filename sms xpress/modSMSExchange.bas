Attribute VB_Name = "modSMSExchange"
Option Explicit
'Public Enum ExcelFileFormat
'    CSV = 1
'    DBF4 = 2
'    Html = 3
'    TextMSDOS = 4
'    TextPrinter = 5
'    TextWindows = 6
'    XMLSpreadsheet = 7
'End Enum
Private ViewHeadings() As String
Private Const LB_DELETESTRING = &H182
Private Const CB_DELETESTRING = &H144
Private Const HWND_TOPMOST = -1
'Private Const TV_FIRST As Long = &H1100
Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH As Long = 260&
Private Const SWP_NOMOVE As Long = 2
Private Const SWP_NOSIZE As Long = 1
Private Const LB_ADDSTRING = &H180
Private Const CB_ADDSTRING = &H143
Private Const flags As Long = SWP_NOMOVE Or SWP_NOSIZE
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
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const CB_FINDSTRINGEXACT = &H158
Private Const LVM_SETITEMCOUNT As Long = 4096 + 47
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetComputerName Lib "Kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
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
Public resp As Integer
Public Intel As New clsIntellisense
Public qrySql As String
'Public Enum Ebt_Type
'    BasEbt = 0
'    SapEbt = 1
'    PerEbt = 2
'End Enum
Public FM As String
Public VM As String
Public NL As String
Public Quote As String
Public RM As String
Public dbExchange As dao.Database
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
    Quote = Chr$(34)
    RM = Chr$(193)
    frmSYS_Splash.Show
    Load frmSMSExchange
    frmSMSExchange.Show
    Unload frmSYS_Splash
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
Public Function RemoveAlpha(ByVal strValue As String) As String
    On Error Resume Next
    strValue = UCase$(strValue)
    strValue = Replace$(strValue, "A", "")
    strValue = Replace$(strValue, "B", "")
    strValue = Replace$(strValue, "C", "")
    strValue = Replace$(strValue, "D", "")
    strValue = Replace$(strValue, "E", "")
    strValue = Replace$(strValue, "F", "")
    strValue = Replace$(strValue, "G", "")
    strValue = Replace$(strValue, "H", "")
    strValue = Replace$(strValue, "I", "")
    strValue = Replace$(strValue, "J", "")
    strValue = Replace$(strValue, "K", "")
    strValue = Replace$(strValue, "L", "")
    strValue = Replace$(strValue, "M", "")
    strValue = Replace$(strValue, "N", "")
    strValue = Replace$(strValue, "O", "")
    strValue = Replace$(strValue, "P", "")
    strValue = Replace$(strValue, "Q", "")
    strValue = Replace$(strValue, "R", "")
    strValue = Replace$(strValue, "S", "")
    strValue = Replace$(strValue, "T", "")
    strValue = Replace$(strValue, "U", "")
    strValue = Replace$(strValue, "V", "")
    strValue = Replace$(strValue, "W", "")
    strValue = Replace$(strValue, "X", "")
    strValue = Replace$(strValue, "Y", "")
    strValue = Replace$(strValue, "Z", "")
    RemoveAlpha = strValue
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
            Rest = StrConcat(Rest, strL, Delim)
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
        sData(tCnt) = StrAdd(Quote, sData(tCnt), Quote)
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
Public Function YyyymmddToNormal(ByVal strValue As String) As String
    On Error Resume Next
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    yyyy = Left$(strValue, 4)
    mm = Mid$(strValue, 5, 2)
    dd = Right$(strValue, 2)
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
            xStr = StrConcat("[", xStr, "]")
        End If
        Select Case lCnt
        Case totArray
            sStr = Concat(sStr, xStr)
        Case Else
            sStr = StrAdd(sStr, xStr, Delim)
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
Public Function DaoFldNames(ByVal dbRs As dao.Recordset, Optional ByVal Delim As String = ",") As String
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
        fL = StrConcat(fL, fN, Delim)
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
    Dim S As String
    N = lstReport.ListItems.Count
    ReDim sTemp(N)
    nDestFile = FreeFile
    Open sDest For Binary Access Write As nDestFile
        For i = 1 To N
            nSrcFile = FreeFile
            S = lstReport.ListItems(i).Text
            Open S For Binary Access Read As nSrcFile
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
    Dim S As String
    N = lstCollection.Count
    ReDim sTemp(N)
    nDestFile = FreeFile
    Open sDest For Binary Access Write As nDestFile
        For i = 1 To N
            nSrcFile = FreeFile
            S = lstCollection(i)
            Open S For Binary Access Read As nSrcFile
                ReDim bTemp(LOF(nSrcFile) - 1)
                Get nSrcFile, , bTemp
                Put nDestFile, , bTemp
            Close nSrcFile
        Next
    Close nDestFile
    ConsolidateFilesCollection = True
    Err.Clear
End Function
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
Public Sub ProgBarInit(progBar As ProgressBar, totItems As Long)
    On Error Resume Next
    ProgBarClose progBar
    progBar.Max = totItems
    progBar.Min = 0
    Err.Clear
End Sub
Public Sub ProgBarClose(progBar As ProgressBar)
    On Error Resume Next
    progBar.Value = 0
    Err.Clear
End Sub
Public Sub dbCreateQueryDef(ByVal Ddb As String, ByVal Qryname As String, ByVal qrySql As String)
    On Error Resume Next
    Dim qdf As dao.QueryDef
    Dim mData As dao.Database
    Set mData = dao.OpenDatabase(Ddb)
    Set qdf = mData.CreateQueryDef(Qryname)
    Select Case Err
    Case 3012           ' already exists
        Set qdf = mData.QueryDefs(Qryname)
    End Select
    qdf.SQL = qrySql
    mData.Close
    Set mData = Nothing
    Set qdf = Nothing
    Err.Clear
End Sub
Public Sub dbDeleteQueries(ByVal DbName As String, ParamArray items())
    On Error Resume Next
    Dim test As String
    Dim Item As Variant
    Dim db As dao.Database
    Set db = dao.OpenDatabase(DbName)
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
Public Function DaoTableExists(ByVal DBase As String, ByVal TbName As String) As Boolean
    On Error Resume Next
    Dim DatCt As Long
    Dim StrDt As String
    Dim zCnt As Long
    Dim db As dao.Database
    TbName = ProperCase(TbName)
    TbName = Iconv(TbName, "t")
    DaoTableExists = False
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoCountRecords(ByVal DBase As String, ByVal Table As String) As Long
    On Error Resume Next
    Dim db As dao.Database
    Dim tb As dao.Recordset
    Set db = dao.OpenDatabase(DBase)
    Set tb = db.OpenRecordset(Table)
    DaoCountRecords = tb.RecordCount
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Function
Function StrConcat(ParamArray items()) As String
    On Error Resume Next
    Dim Item As Variant
    Dim NewString As String
    NewString = ""
    For Each Item In items
        NewString = Concat(NewString, CStr(Item))
    Next
    StrConcat = NewString
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
Public Function StrAdd(ByVal Strdest As String, ByVal Straddstring As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim NewString As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    NewString = Concat(Strdest, Straddstring)
    NewString = Concat(NewString, Delim)
    StrAdd = NewString
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
            strHead = StrAdd(strHead, strName, ",")
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
            StrNew = StrConcat(StrNew, spItems(spItemCnt), Delim)
        End If
    Next
    MvRemoveItems = RemoveDelim(StrNew, Delim)
    Err.Clear
End Function
Public Function DaoTableFieldAutoIncrement(ByVal DBase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As dao.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim att As Integer
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoTableIndexes(ByVal DBase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As dao.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoTablePrimaryIndexes(ByVal DBase As String, ByVal TbName As String) As String '
    On Error Resume Next
    Dim db As dao.Database
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoTableFieldNames(ByVal DBase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As dao.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoTableFieldTypes(ByVal DBase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As dao.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Dim fType As Long
    Set db = dao.OpenDatabase(DBase)
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
Public Function DaoTableFieldSizes(ByVal DBase As String, ByVal TbName As String) As String
    On Error Resume Next
    Dim db As dao.Database
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    Set db = dao.OpenDatabase(DBase)
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
            sLine = StrConcat(sLine, sporiginal(spCnt), Delim)
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
                strR = StrConcat(strR, strData, Delim)
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
                strR = StrConcat(strR, strData, Delim)
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
Function NoCommas(ByVal strValue As String) As String
    On Error Resume Next
    NoCommas = Replace$(strValue, ",", "")
    Err.Clear
End Function
Function ProperAmount(ByVal strValue As String) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsVal As String
    rsStr = ""
    rsTot = Len(strValue)
    For rsCnt = 1 To rsTot
        rsVal = Mid$(strValue, rsCnt, 1)
        If InStr(1, "-.0123456789", rsVal) > 0 Then
            rsStr = rsStr & rsVal
        End If
    Next
    rsStr = Trim$(rsStr)
    If Len(rsStr) = 0 Then rsStr = "0.00"
    If InStr(1, rsStr, ".") = 0 Then rsStr = rsStr & ".00"
    strValue = CDbl(rsStr)
    ProperAmount = Format$(strValue, "###0.00")
    Err.Clear
End Function
Public Function MakeMoney(ByVal strValue As String) As String
    On Error Resume Next
    strValue = ProperAmount(strValue)
    MakeMoney = Format$(strValue, "#,##0.00")
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
Function MvRemoveBlanks(ByVal strValue As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    xRslt = ""
    Call StrParse(xData, strValue, Delim)
    xTot = UBound(xData)
    For xCnt = 1 To xTot
        If Len(Trim$(xData(xCnt))) > 0 Then
            xRslt = StrConcat(xRslt, xData(xCnt), Delim)
        End If
    Next
    xRslt = RemoveDelim(xRslt, Delim)
    MvRemoveBlanks = xRslt
    Err.Clear
End Function
Public Function MvReplaceItem(ByVal strValue As String, ByVal strItem As String, ByVal StrReplaceWith As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spTot As Long
    Dim spCnt As Long
    Call StrParse(spItems, strValue, Delim)
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
        .Item(2).Width = 4005
        .Item(2).Bevel = sbrInset
        .Item(3).Style = 0
        .Item(3).Width = 4005
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
    Dim strSize As Long
    Dim LAST2 As String
    Dim tmpstring As String
    Dim NL As String
    NL = NL = Chr$(13) + Chr$(10)
    tmpstring = StrString
    LAST2 = Right$(tmpstring, 2)
    Do While LAST2 = NL
        strSize = Len(tmpstring) - 2
        tmpstring = Left$(tmpstring, strSize)
        LAST2 = Right$(tmpstring, 2)
    Loop
    RemAllNL = tmpstring
    Err.Clear
End Function
Public Function LstViewUpdate(Arrfields() As String, LstView As ListView, Optional ByVal lstIndex As String = "") As Long
    On Error Resume Next
    Dim ItmX As ListItem
    Dim fldCnt As Integer
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
Function FormatTextForWord(ByVal strValue As String, Optional ByVal NumberOfTabs As Long = 4) As String
    On Error Resume Next
    Dim spTot As Long
    Dim spCnt As Long
    Dim spDat() As String
    Dim sTabs As String
    Dim NL As String
    NL = NL = Chr$(13) + Chr$(10)
    sTabs = String$(NumberOfTabs, vbTab)
    Call StrParse(spDat, strValue, NL)
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
        NewString = StrAdd(NewString, NewItem, Delim)
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
    Dim strSize As Integer
    If Len(sFormat) = 0 Then
        sFormat = "M"
    End If
    Select Case UCase$(sFormat)
    Case "MDY", "DMY", "YDM", "YMD", "DYM", "MYD"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        syy = Right$(spDate(3), 2)
        Select Case UCase$(sFormat)
        Case "MDY":         Oconv = StrConcat(spDate(2), spDate(1), syy)
        Case "DMY":         Oconv = StrConcat(spDate(1), spDate(2), syy)
        Case "YMD":         Oconv = StrConcat(syy, spDate(2), spDate(1))
        Case "YDM":         Oconv = StrConcat(syy, spDate(1), spDate(2))
        Case "DYM":         Oconv = StrConcat(spDate(1), syy, spDate(2))
        Case "MYD":         Oconv = StrConcat(spDate(2), syy, spDate(1))
        End Select
    Case "MDYY", "DMYY", "YYDM", "YYMD", "DYYM", "MYYD"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        Select Case UCase$(sFormat)
        Case "MDYY":         Oconv = StrConcat(spDate(2), spDate(1), spDate(3))
        Case "DMYY":         Oconv = StrConcat(spDate(1), spDate(2), spDate(3))
        Case "YYMD":         Oconv = StrConcat(spDate(3), spDate(2), spDate(1))
        Case "YYDM":         Oconv = StrConcat(spDate(3), spDate(1), spDate(2))
        Case "DYYM":         Oconv = StrConcat(spDate(1), spDate(3), spDate(2))
        Case "MYYD":         Oconv = StrConcat(spDate(2), spDate(3), spDate(1))
        End Select
    Case "YYMM", "YYYYMM"
        theDate = Oconv(sValue, "D")
        Call StrParse(spDate, theDate, "/")
        Oconv = StrConcat(spDate(3), spDate(2))
    Case "F"
        Oconv = Replace$(sValue, "%", "/")
    Case "M"
        If sValue = "." Then
            sValue = "000"
        End If
        sValue = DotAmount(sValue)
        Oconv = Format$(sValue, "#,##0.00")
    Case "D"
        strSize = Len(sValue)
        Select Case strSize
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
    MonthYearDesc = StrAdd(StrMonthName(smm), " ", syy)
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
    YearMonthDesc = StrAdd(syy, " ", StrMonthName(smm))
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
    Dim strSize As Integer
    Select Case Trim$(sAmount)
    Case Is <> ""
        Select Case InStr(sAmount, ".")
        Case 0 ' the amount has no dot
            DotAmount = sAmount & ".00"
            strSize = Len(sAmount)
            Select Case strSize
            Case 1
                sAmount = sAmount & "00"
            Case 2
                s_fpart = Left$(sAmount, 1)
                s_epart = Right$(sAmount, 1)
                Select Case s_fpart
                Case "-": sAmount = StrAdd(s_fpart, "00", s_epart)
                Case Else: sAmount = sAmount & "00"
                End Select
            End Select
            s_size = Len(sAmount)
            s_cents = Right$(sAmount, 2)   ' the last two values
            s_firstpart = s_size - 2
            s_numbers = Left$(sAmount, s_firstpart)
            DotAmount = StrAdd(s_numbers, ".", s_cents)
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
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '" & spStr & "' " & AndOr & " "
            End If
        Case "having"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '" & spStr & "' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '" & spStr & "' " & AndOr & " "
            End If
        Case "like"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '*" & spStr & "*' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " " & StrOperation & " '%" & spStr & "%' " & AndOr & " "
            End If
        Case "likestart", "startwith"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '" & spStr & "*' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '" & spStr & "%' " & AndOr & " "
            End If
        Case "likeend", "endwith"
            If sType = "db" Then
                spStr = SQLQuote(SearchField) & " like '*" & spStr & "' " & AndOr & " "
            Else
                spStr = SQLQuote(SearchField) & " like '%" & spStr & "' " & AndOr & " "
            End If
        End Select
        rslt = StrConcat(rslt, spStr)
NextResult:
    Next
    rslt = RemoveDelim(rslt, AndOr & " ")
    BuildSQL = rslt
    Err.Clear
End Function
Public Function IsPwdValid(ByVal strValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, strValue, "[") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "]") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ".") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "*") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ">") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "<") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ",") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "`") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "#") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "!") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "/") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "\") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "|") > 0), 1, 0)
    intH = intH + IIf(Len(strValue) < 8, 1, 0)
    If intH = 0 Then
        IsPwdValid = True
    Else
        IsPwdValid = False
    End If
    Err.Clear
End Function
Public Function HasSpecial(ByVal strValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, strValue, ".") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "@") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "$") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "%") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "^") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "&") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "(") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ")") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "-") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "}") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "{") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ":") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, ";") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "?") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "~") > 0), 1, 0)
    If intH = 0 Then
        HasSpecial = False
    Else
        HasSpecial = True
    End If
    Err.Clear
End Function
Public Function HasNumber(ByVal strValue As String) As Boolean
    On Error Resume Next
    Dim intH As Integer
    intH = 0
    intH = intH + IIf((InStr(1, strValue, "0") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "1") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "2") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "3") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "4") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "5") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "6") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "7") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "8") > 0), 1, 0)
    intH = intH + IIf((InStr(1, strValue, "9") > 0), 1, 0)
    If intH = 0 Then
        HasNumber = False
    Else
        HasNumber = True
    End If
    Err.Clear
End Function
Public Function IsBlank(ObjectName As Variant, ByVal fldName As String, Optional msgType As Integer = 1) As Boolean
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
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ComboBox Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ImageCombo Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is CheckBox Then
        If ObjectName.Value = 0 Then
            strO = "select"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is ListBox Then
        If (ObjectName.ListCount - 1) = -1 Then
            strO = "select"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is OptionButton Then
        If ObjectName.Value = False Then
            strO = "select"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
            End If
            IsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeOf ObjectName Is Label Then
        If Len(Trim$(ObjectName.Caption)) = 0 Then
            strO = "specify"
            GoSub CompileError
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
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
            If msgType = 1 Then
                MyPrompt strM, "o", "w", ProperCase(StrT)
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
Public Sub LstViewFromComputerFolder(lstReport As ListView, ByVal strFolder As String, ByVal StrHeading As String, Optional ByVal xPattern As String = "*.txt")
    On Error Resume Next
    If Len(strFolder) = 0 Then Exit Sub
    Dim colFiles As New Collection
    lstReport.View = lvwReport
    lstReport.Sorted = True
    LstViewMakeHeadings lstReport, StrHeading
    Set colFiles = MyFilesCollection(strFolder, xPattern)
    LstViewFromCollection lstReport, colFiles
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Public Sub LstViewFromFolder(lstReport As ListView, ByVal strFolder As String, ByVal StrHeading As String, Optional ByVal xPattern As String = "*.txt")
    On Error Resume Next
    If Len(strFolder) = 0 Then Exit Sub
    Dim colFiles As New Collection
    lstReport.View = lvwReport
    lstReport.Sorted = True
    LstViewMakeHeadings lstReport, StrHeading
    Set colFiles = MyFilesCollection(strFolder, xPattern)
    LstViewFromCollection lstReport, colFiles
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Function MyFilesCollection(ByVal strFolder As String, Optional ByVal StrPattern As String = "*.*") As Collection
    On Error Resume Next
    Dim rsTot As Long
    Dim strFile As String
    Dim strP As String
    Dim colNew As New Collection
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFolder As Scripting.Folder
    Dim fsoFile As Scripting.File
    strP = LCase$(MvField(StrPattern, 2, "."))
    Set fsoFolder = fso.GetFolder(strFolder)
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
Public Function ExtractNumbers(ByVal strValue As String) As String
    On Error Resume Next
    Dim i As Long
    Dim sResult As String
    Dim iLen As Long
    Dim myStr As String
    sResult = ""
    iLen = Len(strValue)
    For i = 1 To iLen
        myStr = Mid$(strValue, i, 1)
        If InStr("0123456789", myStr) > 0 Then
            sResult = sResult & myStr
        End If
    Next
    ExtractNumbers = sResult
    Err.Clear
End Function
Function StringPart(ByVal strValue As String, Optional ByVal PartPosition As Long = 1, Optional ByVal Delimiter As String = ",", Optional TrimValue As Boolean = True) As String
    On Error Resume Next
    Dim xResult As String
    Dim xArray() As String
    If Len(strValue) = 0 Then Exit Function
    xArray = Split(strValue, Delimiter)
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
        sData(tCnt) = StrAdd("'", sData(tCnt), "'")
    Next
    MvQuote = MvFromArray(sData, Delim)
    Err.Clear
End Function
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
        StartEndDate = StrAdd(sDate, ",", eDate)
    End Select
    Err.Clear
End Function
Public Function SwapDate(ByVal strValue As String, Optional ConvertMySQL As Boolean = True) As String
    On Error Resume Next
    Dim SY As String
    Dim SM As String
    Dim SD As String
    strValue = Format$(strValue, "dd/mm/yyyy")
    SY = MvField(strValue, 3, "/")
    SM = MvField(strValue, 2, "/")
    SD = MvField(strValue, 1, "/")
    strValue = SM & "/" & SD & "/" & SY
    If ConvertMySQL = True Then
        SwapDate = SY & "-" & SM & "-" & SD
    Else
        SwapDate = strValue
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
Public Function MvIconv(ByVal strData As String, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    MvIconv = Replace$(strData, vbNewLine, Delim)
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
Public Function FileName_Validate(ByVal strValue As String) As String
    On Error Resume Next
    Dim fPath As String
    Dim fFileN As String
    Dim fExt As String
    fPath = FileToken(strValue, "p")
    fFileN = FileToken(strValue, "fo")
    fExt = FileToken(strValue, "e")
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
Public Sub DaoCreateIndexes(ByVal DbName As String, ByVal TbName As String, ByVal IndexNames As String)
    On Error Resume Next
    Dim dbs As dao.Database
    Dim tdf As dao.TableDef
    Dim idxLoop As dao.Index
    Dim IdxName As String
    Dim spIndexes() As String
    Dim spTot As Integer
    Dim spCnt As Integer
    Call StrParse(spIndexes, IndexNames, ",")
    spTot = UBound(spIndexes)
    Set dbs = dao.OpenDatabase(DbName)
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
Function ReadableFigure(ByVal strValue As String) As String
    On Error Resume Next
    strValue = Replace$(strValue, ",", "")
    ReadableFigure = Replace$(strValue, ".", "")
    Err.Clear
End Function
Public Sub DaoTableNamesToLstView(DBase As String, LstView As ListView, Optional ByVal boolClear As Boolean = True)
    On Error Resume Next
    'loads table names into a listview
    Dim db As dao.Database
    Dim StrDt As String
    Dim rsCnt As Long
    Dim rsTot As Long
    ' clear the listview if specified
    If boolClear = True Then LstView.ListItems.Clear
    ' open the database
    Set db = dao.OpenDatabase(DBase)
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
Public Function MvRecord(tb As dao.Recordset, totFlds As Long, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim cntFlds As Long
    Dim rsValues As String
    rsValues = ""
    For cntFlds = 0 To totFlds
        Select Case cntFlds
        Case totFlds
            rsValues = rsValues & tb.Fields(cntFlds).Value & ""
        Case Else
            rsValues = rsValues & tb.Fields(cntFlds).Value & "" & Delim
        End Select
    Next
    MvRecord = rsValues
    Err.Clear
End Function
Public Function Iconv(ByVal sValue As String, Optional ByVal sFormat As String = "") As String
    On Error Resume Next
    Dim sRslt As String
    Dim i As Long
    Dim ch As String
    Dim L As Long
    Dim sN As String
    sRslt = sValue
    Select Case UCase$(sFormat)
    Case ""
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "(", "")
        sRslt = Replace$(sRslt, ")", "")
        sRslt = Replace$(sRslt, "~", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "@", "")
        sRslt = Replace$(sRslt, "#", "")
        sRslt = Replace$(sRslt, "$", "")
        sRslt = Replace$(sRslt, "%", "")
        sRslt = Replace$(sRslt, "^", "")
        sRslt = Replace$(sRslt, "&", "")
        sRslt = Replace$(sRslt, "*", "")
        sRslt = Replace$(sRslt, "_", "")
        sRslt = Replace$(sRslt, "-", "")
        sRslt = Replace$(sRslt, "=", "")
        sRslt = Replace$(sRslt, "|", "")
        sRslt = Replace$(sRslt, "\", "")
        sRslt = Replace$(sRslt, ":", "")
        sRslt = Replace$(sRslt, ";", "")
        sRslt = Replace$(sRslt, "<", "")
        sRslt = Replace$(sRslt, ">", "")
        sRslt = Replace$(sRslt, "?", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "+", "")
        sRslt = Replace$(sRslt, "{", "")
        sRslt = Replace$(sRslt, "}", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, Quote, "")
    Case "Q"
        sRslt = Replace$(sRslt, "''", "")
        sRslt = Replace$(sRslt, "'", "")
    Case "F"
        sRslt = Replace$(sRslt, "/", "%")
        sRslt = Replace$(sRslt, "\", "%")
        sRslt = Replace$(sRslt, "|", "%")
    Case "C"
        sRslt = Replace$(sRslt, ",", "")
    Case "M"
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, ".", "")
    Case "S"
        L = Len(sRslt)
        sRslt = sRslt
        If L = 0 Then
            Err.Clear
            Exit Function
        End If
        sN = ""
        For i = 1 To L
            ch = Mid$(sRslt, i, 1)
            If ch = " " Then
                sN = sN & ch
            End If
            If ch >= "a" Then
                If ch <= "z" Then
                    sN = sN & ch
                End If
            End If
            If ch >= "A" Then
                If ch <= "Z" Then
                    sN = sN & ch
                End If
            End If
        Next
        sRslt = sN
    Case "T"
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, Quote, "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, ",", "")
    End Select
    Iconv = sRslt
    Err.Clear
End Function
Public Function SQLQuote(ByVal strData As String, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    ' inserts a bracket to strings to use with mysqk
    Dim sData() As String
    Dim tCnt As Integer
    Dim wCnt As Integer
    Dim rslt As String
    rslt = ""
    sData = Split(strData, Delim)
    wCnt = UBound(sData)
    For tCnt = 0 To wCnt
        sData(tCnt) = "[" & sData(tCnt) & "]"
        If tCnt = wCnt Then
            rslt = rslt & sData(tCnt)
        Else
            rslt = rslt & sData(tCnt) & Delim
        End If
    Next
    SQLQuote = rslt
    Err.Clear
End Function
Public Sub CboBoxLoadKeys(OpenDb As dao.Database, ByVal RecordSource As String, ByVal Datafields As String, cboList As Variant, Optional ByVal Delim As String = "", Optional ByVal cboClear As String = "", Optional ByVal IndexFld As String = "", Optional MakeProperCase As Boolean = False, Optional RemoveDuplicates As Boolean = False)
    On Error Resume Next
    Dim dtFlds() As String
    Dim strLine As String
    Dim dtCnt As Integer
    Dim ADOR As dao.Recordset
    Dim fldName As String
    Dim fldValue As String
    Dim wCnt As Long
    Dim idxData As String
    Dim rsCnt As Long
    Dim rsTot As Long
    If Len(cboClear) = 0 Then cboList.Clear
    If Len(Delim) = 0 Then Delim = VM
    Call StrParse(dtFlds, Datafields, ",")
    wCnt = UBound(dtFlds)
    If LCase$(Left$(RecordSource, 6)) = "select" Then
        Set ADOR = OpenDb.OpenRecordset(RecordSource)
    Else
        Set ADOR = OpenDb.OpenRecordset("select distinct " & Datafields & " from [" & RecordSource & "] order by " & Datafields)
    End If
    ADOR.MoveLast
    rsTot = ADOR.RecordCount
    ADOR.MoveFirst
    For rsCnt = 1 To rsTot
        strLine = ""
        For dtCnt = 1 To wCnt
            fldName = dtFlds(dtCnt)
            fldValue = ADOR.Fields(fldName)
            If dtCnt = wCnt Then
                strLine = strLine & fldValue
            Else
                strLine = strLine & fldValue & Delim
            End If
        Next
        If MakeProperCase = True Then strLine = ProperCase(strLine)
        If RemoveDuplicates = True Then
            LstBoxUpdate cboList, strLine
        Else
            cboList.AddItem strLine
            If Len(IndexFld) > 0 Then
                idxData = ADOR.Fields(IndexFld)
                cboList.ItemData(cboList.NewIndex) = idxData
            End If
        End If
        ADOR.MoveNext
    Next
    ADOR.Close
closeLoad:
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
        Delim = Chr$(253)
    End If
    lstTot = LstView.ListItems.Count
    For lstCnt = 1 To lstTot
        bOp = LstView.ListItems(lstCnt).Checked
        Select Case bOp
        Case True
            lstStr = LstViewGetRow(LstView, lstCnt)
            retStr = retStr & lstStr(lngPos) & Delim
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
    LstViewCheckedToMV = retStr
    Err.Clear
End Function
