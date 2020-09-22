VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntellisense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_txtComplete As TextBox
Attribute m_txtComplete.VB_VarHelpID = -1
Private m_strDelimeter As String
Private m_strList As String
Private intelFile As String
Private m_Path As String
Private Property Set CompleteTextbox(ByRef txt As TextBox)
    On Error Resume Next
    Set m_txtComplete = txt
    If DirExists(SensePath & "\Intellisense") = False Then MkDir SensePath & "\Intellisense"
    intelFile = SensePath & "\Intellisense\" & m_txtComplete.Name & ".sense"
    Err.Clear
End Property
Private Property Get SearchList() As String
    On Error Resume Next
    SearchList = m_strList
    Err.Clear
End Property
Private Property Let SearchList(ByVal str As String)
    On Error Resume Next
    m_strList = str
    Err.Clear
End Property
Private Property Get SensePath() As String
    On Error Resume Next
    SensePath = m_Path
    Err.Clear
End Property
Private Property Let SensePath(ByVal str As String)
    On Error Resume Next
    m_Path = str
    Err.Clear
End Property
Private Property Get Delimeter() As String
    On Error Resume Next
    Delimeter = m_strDelimeter
    Err.Clear
End Property
Private Property Let Delimeter(ByVal str As String)
    On Error Resume Next
    m_strDelimeter = str
    Err.Clear
End Property
Private Sub AddToIsense()
    On Error Resume Next
    Dim NewValue As String
    If Len(Trim$(m_txtComplete.Text)) = 0 Then Exit Sub
    If m_txtComplete.MultiLine = False Then
        StringRemoveNewLines
    End If
    If MvSearch(SearchList, m_txtComplete.Text, Delimeter) = 0 Then
        NewValue = Delimeter & m_txtComplete.Text
        SearchList = SearchList & NewValue
    End If
    SearchList = Replace$(SearchList, vbNewLine, "")
    SearchListDataUpdate
    Err.Clear
End Sub
Private Sub StringRemoveNewLines()
    On Error Resume Next
    Dim strSize As Long
    Dim LAST2 As String
    LAST2 = Right$(m_txtComplete.Text, 2)
    Do While LAST2 = vbNewLine
        strSize = Len(m_txtComplete.Text) - 2
        m_txtComplete.Text = Left$(m_txtComplete.Text, strSize)
        LAST2 = Right$(m_txtComplete.Text, 2)
    Loop
    m_txtComplete.SelStart = Len(m_txtComplete.Text)
    Err.Clear
End Sub
Private Function SearchListData() As String
    On Error Resume Next
    Dim sLen As Long
    Dim myBuf As String
    Dim fileNum As String
    Dim Size As Long
    SearchListData = ""
    If FileExists(intelFile) = True Then
        fileNum = FreeFile
        Size = FileLen(intelFile)
        myBuf = String$(Size, "*")
        Open intelFile For Input Access Read As #fileNum
            sLen = LOF(fileNum)
            SearchListData = Input(sLen, #fileNum)
            SearchListData = Replace$(SearchListData, vbNewLine, "")
        Close #fileNum
    End If
    Err.Clear
End Function
Private Sub SearchListDataUpdate()
    On Error Resume Next
    Dim iFileNum As Integer
    iFileNum = FreeFile
    Open intelFile For Output As #iFileNum
        Print #iFileNum, Replace$(SearchList, vbNewLine, "")
    Close #iFileNum
    Err.Clear
End Sub
Public Sub Sense(ByVal Path As String, objTextBox As TextBox, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    SensePath = Path
    Set CompleteTextbox = objTextBox
    m_txtComplete_KeyUp KeyCode, Shift
    Err.Clear
End Sub
Private Sub Class_Initialize()
    On Error Resume Next
    Delimeter = Chr$(254)
    Err.Clear
End Sub
Private Sub m_txtComplete_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim strSearchText As String
    Dim intLength As Integer
    Dim varArray As Variant
    Dim i_Tot As Integer
    SearchList = SearchListData
    With m_txtComplete
        If KeyCode <> vbKeyBack And KeyCode > 48 Then
            If InStr(1, m_strList, .Text, vbTextCompare) <> 0 Then
                varArray = Split(m_strList, m_strDelimeter)
                i_Tot = UBound(varArray)
                For i = 0 To i_Tot
                    strSearchText = Trim$(varArray(i))
                    If InStr(1, strSearchText, .Text, vbTextCompare) And (Left$(.Text, 1) = Left$(strSearchText, 1)) And .Text <> "" Then
                        .SelText = ""
                        .SelLength = 0
                        intLength = Len(.Text)
                        .Text = .Text & Right$(strSearchText, Len(strSearchText) - Len(.Text))
                        .SelStart = intLength
                        .SelLength = Len(.Text)
                        Err.Clear
                        Exit Sub
                    End If
                Next
            End If
        End If
    End With
    Err.Clear
End Sub
Private Sub m_txtComplete_Validate(Cancel As Boolean)
    On Error Resume Next
    AddToIsense
    Err.Clear
End Sub
Private Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = "", Optional TrimItems As Boolean = False) As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = 0
    If Len(StringMv) = 0 Then
        MvSearch = 0
        Err.Clear
        Exit Function
    End If
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(TheFields, StringMv, Delim)
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
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
    Err.Clear
End Function
Private Function StrParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = "") As Long
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
    ReDim retarray(VarE + 1)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
