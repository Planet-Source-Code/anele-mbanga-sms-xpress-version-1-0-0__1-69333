VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl GSM 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GSM.ctx":0000
   ScaleHeight     =   660
   ScaleWidth      =   930
   ToolboxBitmap   =   "GSM.ctx":0822
   Begin MSCommLib.MSComm MSComm 
      Left            =   240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "GSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mSpeed As String
Event Response(ByVal Result As String)
Private mCommPort As Integer
Private mPortOpen As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private m_Buffer As String
Private mLogFile As String
Private mReadDelete As String
Private mWriteSend As String
Private mReceived As String
Private mReceivedUsed As Long
Private mReceivedCapacity As Long
Private mReadDeleteUsed As Long
Private mReadDeleteCapacity As Long
Private mWriteSendUsed As String
Private mWriteSendCapacity As String
Private mpbMemory As String
Private mpbCapacity As Long
Private mpbUsed As Long
Public Enum SmsMemoryStorageEnum
    SimMemory = 0
    MobileEquipmentMemory = 1
    BothMemories = 2
    ReadMemorySetting = 3
    BroadcastMessageStorage = 4
    StatusReportStorage = 5
    TerminalAdapterStorage = 6
End Enum
Public Enum MessageFormatEnum
    TextFormat = 1
    PDUFormat = 0
    ReadFormat = 2
End Enum
Public Enum EchoEnum
    EchoOff = 0
    EchoOn = 1
End Enum
Public Enum PhoneBookMemoryStorageEnum
    SimPhoneBook = 0
    MobileEquipmentPhoneBook = 1
    BothPhoneBooks = 2
    ReadPhoneBookSetting = 3
End Enum
Public Enum CentreNumberEnum
    SetCentreNumber = 0
    ReadCentreNumber = 1
End Enum
Public Enum SmsTypesEnum
    RecRead = 0
    RecUnread = 1
    Rec = 2
    StoSent = 3
    StoUnsent = 4
    All = 5
End Enum
Public Property Get CommPort() As Integer
    On Error Resume Next
    ' get the port used by the modem
    CommPort = mCommPort
    Err.Clear
End Property
Public Property Let CommPort(nCommPort As Integer)
    On Error Resume Next
    ' set the port to use for the modem connection
    MSComm.CommPort = nCommPort
    mCommPort = nCommPort
    PropertyChanged "CommPort"
    Err.Clear
End Property
Private Function FixSmsDate(sDate As String) As String
    On Error Resume Next
    Dim syyyy As String
    Dim smm As String
    Dim sdd As String
    syyyy = MvField(sDate, 1, "/")
    smm = MvField(sDate, 2, "/")
    sdd = MvField(sDate, 3, "/")
    sdd = sdd & "/" & smm & "/" & syyyy
    FixSmsDate = Format$(sdd, "dd/mm/yyyy")
    Err.Clear
End Function
Public Property Get LogFile() As String
    On Error Resume Next
    ' get the path of the log file
    LogFile = mLogFile
    Err.Clear
End Property
Public Property Let LogFile(nLogFile As String)
    On Error Resume Next
    ' set the path of the log file
    mLogFile = nLogFile
    PropertyChanged "LogFile"
    Err.Clear
End Property
Public Property Get Settings() As String
    On Error Resume Next
    ' get the settings of the modem
    Settings = MSComm.Settings
    Err.Clear
End Property
Public Property Get VM() As String
    On Error Resume Next
    ' value delimiter
    VM = Chr$(253)
    Err.Clear
End Property
Public Property Get FM() As String
    On Error Resume Next
    ' value delimiter
    FM = Chr$(254)
    Err.Clear
End Property
Public Property Get Quote() As String
    On Error Resume Next
    ' value double quote
    Quote = Chr$(34)
    Err.Clear
End Property
Public Property Get Speed() As String
    On Error Resume Next
    ' get the speed of the modem
    Speed = mSpeed
    Err.Clear
End Property
Public Property Let Speed(nSpeed As String)
    On Error Resume Next
    ' set the speed of the modem, this will change settings
    mSpeed = nSpeed
    MSComm.Settings = nSpeed & ",n,8,1"
    PropertyChanged "Speed"
    PropertyChanged "Settings"
    Err.Clear
End Property
Public Property Get PortOpen() As Boolean
    On Error Resume Next
    ' get the status of the port
    PortOpen = mPortOpen
    Err.Clear
End Property
Public Property Let PortOpen(nPortOpen As Boolean)
    On Error GoTo ErrMsg
    ' set the status of the port
    MSComm.PortOpen = nPortOpen
    mPortOpen = nPortOpen
    PropertyChanged "PortOpen"
    Err.Clear
    Exit Property
ErrMsg:
    RaiseEvent Response(Err.Description)
    Err.Clear
End Property
Private Sub MSComm_OnComm()
    On Error Resume Next
    ' read contents of the data sent by modem
    Select Case MSComm.CommEvent
    Case comEvReceive
        m_Buffer = MSComm.Input
    End Select
    Err.Clear
End Sub
Public Function Connect(MyPort As String, MySpeed As String) As String
    On Error Resume Next
    ' connect to the gsm modem by specifying the port number and speed of the modem
    If MSComm.PortOpen = True Then MSComm.PortOpen = False
    Me.CommPort = Val(MyPort)
    Me.Speed = MySpeed
    MSComm.DTREnable = True
    MSComm.RTSEnable = True
    MSComm.RThreshold = 1
    MSComm.InBufferSize = 1024
    Me.PortOpen = True
    Me.Echo EchoOn
    Connect = Me.IsReady
    Err.Clear
End Function
Public Function IsReady() As String
    On Error Resume Next
    ' is the modem ready, get the status
    IsReady = Request("AT", , 2, vbCrLf)
    RaiseEvent Response(IsReady)
    Err.Clear
End Function
Public Property Get ManufacturerIdentification() As String
    On Error Resume Next
    ' get the gadget manufacturer identification
    ManufacturerIdentification = Request("AT+GMI", , 2, vbCrLf)
    RaiseEvent Response(ManufacturerIdentification)
    Err.Clear
End Property
Public Property Get ModemSerialNumber() As String
    On Error Resume Next
    ' get the gadget modem serial/imei number
    ModemSerialNumber = Request("AT+GSN", , 2, vbCrLf)
    RaiseEvent Response(ModemSerialNumber)
    Err.Clear
End Property
Public Property Get RevisionIdentification() As String
    On Error Resume Next
    ' get the gadget revision identification
    Dim mAnswer As String
    mAnswer = Request("AT+GMR", , 2, vbCrLf)
    RevisionIdentification = MvField(mAnswer, 2, ":")
    RaiseEvent Response(RevisionIdentification)
    Err.Clear
End Property
Public Property Get ModelIdentification() As String
    On Error Resume Next
    ' get the gadget model indentification number
    ModelIdentification = Request("AT+GMM", , 2, vbCrLf)
    RaiseEvent Response(ModelIdentification)
    Err.Clear
End Property
Public Function SMS_NewMessageIndicate(bSet As Boolean) As String
    On Error Resume Next
    ' tell the gadget to notify computer of new sms received
    If bSet = True Then
        SMS_NewMessageIndicate = Request("AT+CNMI=1,1,0,0,0", , 2, vbCrLf)
    Else
        SMS_NewMessageIndicate = Request("AT+CNMI=0,0,0,0,0", , 2, vbCrLf)
    End If
    RaiseEvent Response(SMS_NewMessageIndicate)
    Err.Clear
End Function
Public Property Get SignalQualityMeasure() As String
    On Error Resume Next
    ' read the signal of the phone and return it as a percentage
    ' the maximum signal is 31
    Dim mAnswer As String
    mAnswer = Request("AT+CSQ", , 2, vbCrLf)
    mAnswer = MvField(mAnswer, 2, ":")
    mAnswer = MvField(mAnswer, 1, ",")
    mAnswer = (Val(mAnswer) / 31) * 100
    SignalQualityMeasure = Round(mAnswer, 0)
    RaiseEvent Response(SignalQualityMeasure)
    Err.Clear
End Property
Public Function PhoneBook_MemoryStorage(PhoneBookSelect As PhoneBookMemoryStorageEnum) As String
    On Error Resume Next
    ' select gadget phonebook memory
    Dim mAnswer As String
    Select Case PhoneBookSelect
    Case SimPhoneBook
        mAnswer = Request("AT+CPBS=" & Me.Quote & "SM" & Me.Quote, "", 2, vbCrLf)
    Case MobileEquipmentPhoneBook
        mAnswer = Request("AT+CPBS=" & Me.Quote & "ME" & Me.Quote, "", 2, vbCrLf)
    Case BothPhoneBooks
        mAnswer = Request("AT+CPBS=" & Me.Quote & "MT" & Me.Quote, "", 2, vbCrLf)
    Case ReadPhoneBookSetting
        mAnswer = Request("AT+CPBS?", , 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mpbCapacity = MvField(mAnswer, 3, ",")
        mpbUsed = MvField(mAnswer, 2, ",")
        mpbMemory = MvField(mAnswer, 1, ",")
    End Select
    PhoneBook_MemoryStorage = mAnswer
    RaiseEvent Response(PhoneBook_MemoryStorage)
    Err.Clear
End Function
Public Function PhoneBook_ReadEntry(Location As Long) As String
    On Error Resume Next
    ' read a specified phonebook entry using a location
    ' an empty location returns blank
    Dim mAnswer As String
    Dim mLoc As String
    Dim mNum As String
    Dim mNam As String
    mAnswer = Request("AT+CPBR=" & Location, , 2, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBR:", ""))
    mLoc = MvField(mAnswer, 1, ",")
    mNum = MvField(mAnswer, 2, ",")
    mNam = MvField(mAnswer, 4, ",")
    mAnswer = mLoc & "," & mNum & "," & mNam
    Select Case mAnswer
    Case ",,", "Not found,Not found,Not found"
        PhoneBook_ReadEntry = ""
    Case Else
        PhoneBook_ReadEntry = mLoc & "," & mNum & "," & mNam
    End Select
    RaiseEvent Response(PhoneBook_ReadEntry)
    Err.Clear
End Function
Public Function PhoneBook_FindEntry(FullName As String) As String
    On Error Resume Next
    ' search for a phonebook entry and return result
    Dim mAnswer As String
    mAnswer = Request("AT+CPBF=" & Me.Quote & FullName & Me.Quote, , 2, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBF:", ""))
    PhoneBook_FindEntry = mAnswer
    RaiseEvent Response(PhoneBook_FindEntry)
    Err.Clear
End Function
Public Function PhoneBook_EntryExists(ByVal FullName As String, ByVal CellNo As String) As String
    On Error Resume Next
    ' search for a phonebook entry for the name and cellphone and return the location
    ' a zero location means not found
    Dim mAnswer As String
    Dim mPhone As String
    Dim mIndex As String
    mAnswer = Request("AT+CPBF=" & Me.Quote & FullName & Me.Quote, , 2, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBF:", ""))
    Select Case mAnswer
    Case "Not Found", ""
        PhoneBook_EntryExists = "0"
    Case Else
        mIndex = MvField(mAnswer, 1, ",")
        mPhone = MvField(mAnswer, 2, ",")
        If mPhone = CellNo Then
            PhoneBook_EntryExists = mIndex
        Else
            PhoneBook_EntryExists = "0"
        End If
    End Select
    RaiseEvent Response(PhoneBook_EntryExists)
    Err.Clear
End Function
Public Function PhoneBook_WriteEntry(Location As Long, ByVal CellNo As String, ByVal FullName As String) As String
    On Error Resume Next
    ' write a phonebook entry at a particular location and refresh the count
    Dim mAnswer As String
    mAnswer = "AT+CPBW=" & Location & "," & Me.Quote & CellNo & Me.Quote & ",129," & Me.Quote & FullName & Me.Quote
    mAnswer = Request(mAnswer, , 2, vbCrLf)
    PhoneBook_WriteEntry = mAnswer
    If mAnswer = "OK" Then Me.PhoneBook_MemoryStorage (ReadPhoneBookSetting)
    RaiseEvent Response(PhoneBook_WriteEntry)
    Err.Clear
End Function
Public Function PhoneBook_DeleteEntry(Location As Long) As String
    On Error Resume Next
    ' delete a phonebook entry using the location
    Dim mAnswer As String
    mAnswer = Request("AT+CPBW=" & Location, , 2, vbCrLf)
    PhoneBook_DeleteEntry = mAnswer
    RaiseEvent Response(PhoneBook_DeleteEntry)
    Err.Clear
End Function
Public Property Get PhoneBook_AvailableIndexes() As String
    On Error Resume Next
    ' get the available phonebook indexes
    Dim mAnswer As String
    mAnswer = Request("AT+CPBR=?", , 2, vbCrLf)
    mAnswer = MvField(mAnswer, 2, ":")
    PhoneBook_AvailableIndexes = mAnswer
    RaiseEvent Response(PhoneBook_AvailableIndexes)
    Err.Clear
End Property
Public Function Echo(EchoStatus As EchoEnum) As Boolean
    On Error Resume Next
    ' turn echo off/on, off results in less traffic
    ' echo off returns the command with the result
    Dim mAnswer As String
    Select Case EchoStatus
    Case EchoOff
        mAnswer = Request("ATE0", , 2, vbCrLf)
    Case EchoOn
        mAnswer = Request("ATE1", , 2, vbCrLf)
    End Select
    If mAnswer = "OK" Then
        Echo = True
    Else
        Echo = False
    End If
    RaiseEvent Response(Echo)
    Err.Clear
End Function
Public Function Request(ByVal strCommand As String, Optional ByVal ExpectedResult As String = "", Optional Position As Long = -1, Optional ControlChars As String = vbCrLf) As String
    On Error GoTo ErrMsg
    ' send a request to the comm port and wait for the result
    ' the result will be delimited by chr(253)
    Dim sResult As String
    If Len(LogFile) > 0 Then
        FileUpdate LogFile, Now() & ", request received " & strCommand, "a"
    End If
    MSComm.Output = strCommand & ControlChars
    sResult = WaitReply(10, ExpectedResult)
    sResult = Replace$(sResult, vbNewLine, VM)
    If Len(LogFile) > 0 Then
        FileUpdate LogFile, Now() & ", response received " & sResult, "a"
    End If
    If Position > 0 Then
        sResult = MvField(sResult, Position, VM)
    Else
        sResult = sResult
    End If
    Request = DescriptiveError(sResult)
    Err.Clear
    Exit Function
ErrMsg:
    RaiseEvent Response(Err.Description)
    Err.Clear
End Function
Private Function WaitReply(lDelay As Long, WaitString As String) As String
    On Error Resume Next
    ' wait process for a data request from the
    ' ms comm port to be finalized
    Dim x As Long
    Dim bOK As Boolean
    DoEvents
    If WaitString = "" Then
        bOK = True
        For x = 1 To lDelay
            Sleep 500
            DoEvents
        Next
    Else
        bOK = False
        For x = 1 To lDelay
            DoEvents
            If InStr(m_Buffer, WaitString) Then
                bOK = True
                Exit For
            Else
                DoEvents
                Sleep 500
                DoEvents
            End If
        Next
    End If
    If Len(m_Buffer) > 0 Then
        DoEvents
    End If
    WaitReply = m_Buffer
    Err.Clear
End Function
Private Function MvField(ByVal strValue As String, Optional ByVal PartPosition As Long = 1, Optional ByVal Delimiter As String = ",", Optional TrimValue As Boolean = True) As String
    On Error Resume Next
    ' return a substring of a string delimited by a string specified
    Dim xResult As String
    Dim xArray() As String
    Dim xSize As Long
    If Len(strValue) = 0 Then Exit Function
    If InStr(1, strValue, Delimiter) = 0 Then
        MvField = strValue
        Err.Clear
        Exit Function
    End If
    xArray = Split(strValue, Delimiter)
    Select Case PartPosition
    Case -1
        PartPosition = UBound(xArray) + 1
    Case 0
        PartPosition = 1
    End Select
    xSize = UBound(xArray)
    If xSize = 0 Then
        MvField = ""
    Else
        xResult = xArray(PartPosition - 1)
        If TrimValue = True Then
            xResult = Trim$(xResult)
        End If
        MvField = xResult
    End If
    Err.Clear
End Function
Private Sub FileUpdate(ByVal filName As String, ByVal filLines As String, Optional ByVal Wora As String = "write")
    On Error Resume Next
    ' update contents of a file by either appending / creating a new entry
    Dim iFileNum As Long
    Dim cDir As String
    cDir = FileToken(filName, "p")
    CreateNestedDirectory cDir
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
Private Function DescriptiveError(ByVal sResult As String) As String
    On Error Resume Next
    ' return descriptive phone error code
    Select Case sResult
    Case "+CME ERROR: 0"
        DescriptiveError = "Phone failure"
    Case "+CME ERROR: 1"
        DescriptiveError = "No connection to phone"
    Case "+CME ERROR: 2"
        DescriptiveError = "Phone adapter link reserved"
    Case "+CME ERROR: 3"
        DescriptiveError = "Operation not allowed"
    Case "+CME ERROR: 4"
        DescriptiveError = "Operation not supported"
    Case "+CME ERROR: 5"
        DescriptiveError = "PH_SIM PIN required"
    Case "+CME ERROR: 6"
        DescriptiveError = "PH_FSIM PIN required"
    Case "+CME ERROR: 7"
        DescriptiveError = "PH_FSIM PUK required"
    Case "+CME ERROR: 10"
        DescriptiveError = "SIM not inserted"
    Case "+CME ERROR: 11"
        DescriptiveError = "SIM PIN required"
    Case "+CME ERROR: 12"
        DescriptiveError = "SIM PUK required"
    Case "+CME ERROR: 13"
        DescriptiveError = "SIM failure"
    Case "+CME ERROR: 14"
        DescriptiveError = "SIM busy"
    Case "+CME ERROR: 15"
        DescriptiveError = "SIM wrong"
    Case "+CME ERROR: 16"
        DescriptiveError = "Incorrect password"
    Case "+CME ERROR: 17"
        DescriptiveError = "SIM PIN2 required"
    Case "+CME ERROR: 18"
        DescriptiveError = "SIM PUK2 required"
    Case "+CME ERROR: 20"
        DescriptiveError = "Memory full"
    Case "+CME ERROR: 21"
        DescriptiveError = "Invalid index"
    Case "+CME ERROR: 22"
        DescriptiveError = "Not found"
    Case "+CME ERROR: 23"
        DescriptiveError = "Memory failure"
    Case "+CME ERROR: 24"
        DescriptiveError = "Text string too long"
    Case "+CME ERROR: 25"
        DescriptiveError = "Invalid characters in text string"
    Case "+CME ERROR: 26"
        DescriptiveError = "Dial string too long"
    Case "+CME ERROR: 27"
        DescriptiveError = "Invalid characters in dial string"
    Case "+CME ERROR: 30"
        DescriptiveError = "No network service"
    Case "+CME ERROR: 31"
        DescriptiveError = "Network timeout"
    Case "+CME ERROR: 32"
        DescriptiveError = "Network not allowed, emergency calls only"
    Case "+CME ERROR: 40"
        DescriptiveError = "Network personalization PIN required"
    Case "+CME ERROR: 41"
        DescriptiveError = "Network personalization PUK required"
    Case "+CME ERROR: 42"
        DescriptiveError = "Network subset personalization PIN required"
    Case "+CME ERROR: 43"
        DescriptiveError = "Network subset personalization PUK required"
    Case "+CME ERROR: 44"
        DescriptiveError = "Service provider personalization PIN required"
    Case "+CME ERROR: 45"
        DescriptiveError = "Service provider personalization PUK required"
    Case "+CME ERROR: 46"
        DescriptiveError = "Corporate personalization PIN required"
    Case "+CME ERROR: 47"
        DescriptiveError = "Corporate personalization PUK required"
    Case "+CME ERROR: 48"
        DescriptiveError = "PH-SIM PUK required"
    Case "+CME ERROR: 100"
        DescriptiveError = "Unknown error"
    Case "+CME ERROR: 103"
        DescriptiveError = "Illegal MS"
    Case "+CME ERROR: 106"
        DescriptiveError = "Illegal ME"
    Case "+CME ERROR: 107"
        DescriptiveError = "GPRS services not allowed"
    Case "+CME ERROR: 111"
        DescriptiveError = "PLMN not allowed"
    Case "+CME ERROR: 112"
        DescriptiveError = "Location area not allowed"
    Case "+CME ERROR: 113"
        DescriptiveError = "Roaming not allowed in this location area"
    Case "+CME ERROR: 126"
        DescriptiveError = "Operation temporary not allowed"
    Case "+CME ERROR: 132"
        DescriptiveError = "Service operation not supported"
    Case "+CME ERROR: 133"
        DescriptiveError = "Requested service option not subscribed"
    Case "+CME ERROR: 134"
        DescriptiveError = "Service option temporary out of order"
    Case "+CME ERROR: 148"
        DescriptiveError = "Unspecified GPRS error"
    Case "+CME ERROR: 149"
        DescriptiveError = "PDP authentication failure"
    Case "+CME ERROR: 150"
        DescriptiveError = "Invalid mobile class"
    Case "+CME ERROR: 256"
        DescriptiveError = "Operation temporarily not allowed"
    Case "+CME ERROR: 257"
        DescriptiveError = "Call barred"
    Case "+CME ERROR: 258"
        DescriptiveError = "Phone is busy"
    Case "+CME ERROR: 259"
        DescriptiveError = "User abort"
    Case "+CME ERROR: 260"
        DescriptiveError = "Invalid dial string"
    Case "+CME ERROR: 261"
        DescriptiveError = "SS not executed"
    Case "+CME ERROR: 262"
        DescriptiveError = "SIM Blocked"
    Case "+CME ERROR: 263"
        DescriptiveError = "Invalid block"
    Case "+CME ERROR: 772"
        DescriptiveError = "SIM powered down"
    Case Else
        DescriptiveError = sResult
    End Select
    Err.Clear
End Function
Public Sub PhoneBook_ListView(progBar As ProgressBar, LstView As ListView, used As Long, capacity As Long, Optional mIcon As String = "", Optional mSmallIcon As String = "")
    On Error Resume Next
    ' load contents of the selected phonebook
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    Dim lstItem As ListItem
    Dim spLine() As String
    Dim lstTotal As Long
    progBar.Max = capacity
    progBar.Min = 0
    progBar.Value = 0
    ' create headings
    LstViewMakeHeadings LstView, "Index,Cellphone No,Full Name"
    ' loop through the phonebook starting from location 1 to the full capacity of the
    ' phone
    For rsCnt = 1 To capacity
        progBar.Value = rsCnt
        ' read entry at specified index
        phEntry = Me.PhoneBook_ReadEntry(rsCnt)
        ' if successfull return index,cellno,fullname
        If Len(phEntry) > 0 Then
            spLine = Split(phEntry, ",")
            Set lstItem = LstView.ListItems.Add(, , spLine(0))
            lstItem.SubItems(1) = spLine(1)
            lstItem.SubItems(2) = spLine(2)
            If Len(mIcon) > 0 Then lstItem.Icon = mIcon
            If Len(mSmallIcon) > 0 Then lstItem.SmallIcon = mSmallIcon
        End If
        DoEvents
        ' ensure that if we have reached the used limit, exit the loop
        ' we do not want to read the empty contacts anymore
        lstTotal = LstView.ListItems.Count
        If lstTotal = used Then Exit For
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub LstViewMakeHeadings(LstView As ListView, ByVal strHeads As String)
    On Error Resume Next
    ' used to create columns in a listview
    Dim fldCnt As Integer
    Dim FldHead() As String
    Dim fldTot As Integer
    Dim colX As Variant
    FldHead = Split(strHeads, ",")
    fldTot = UBound(FldHead)
    LstView.ColumnHeaders.Clear
    LstView.ListItems.Clear
    LstView.Sorted = False
    ' first column should be left aligned
    Set colX = LstView.ColumnHeaders.Add(, , FldHead(0), 1440)
    For fldCnt = 1 To fldTot
        Set colX = LstView.ColumnHeaders.Add(, , FldHead(fldCnt), 1440)
    Next
    LstView.View = lvwReport
    LstView.Checkboxes = True
    LstView.GridLines = True
    LstView.FullRowSelect = True
    LstView.Refresh
    Err.Clear
End Sub
Public Property Get PhoneBook_AvailableIndex() As Long
    On Error Resume Next
    ' get next available index
    Dim rsCnt As Long
    Dim phEntry As String
    Dim pCapacity As Long
    Call Me.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
    pCapacity = Me.PhoneBook_Capacity
    PhoneBook_AvailableIndex = -1
    For rsCnt = 1 To pCapacity
        ' read entry at specified index
        phEntry = Me.PhoneBook_ReadEntry(rsCnt)
        ' if successfull return index,cellno,fullname
        If Len(phEntry) = 0 Then
            PhoneBook_AvailableIndex = rsCnt
            Exit For
        End If
        DoEvents
    Next
    RaiseEvent Response(PhoneBook_AvailableIndex)
    Err.Clear
End Property
Public Function PhoneBook_AddEntry(ByVal sNumber As String, ByVal sName As String) As String
    On Error Resume Next
    'add an entry to the phonebook
    Dim availableIndex As Long
    availableIndex = Me.PhoneBook_AvailableIndex
    If availableIndex = -1 Then
        PhoneBook_AddEntry = "Phonebook Full"
    Else
        PhoneBook_AddEntry = Me.PhoneBook_WriteEntry(availableIndex, sNumber, sName)
    End If
    RaiseEvent Response(PhoneBook_AddEntry)
    Err.Clear
End Function
Public Function SMS_MessageFormat(MessageFormatAction As MessageFormatEnum) As String
    On Error Resume Next
    ' message format management
    Dim mAnswer As String
    Select Case MessageFormatAction
    Case TextFormat
        SMS_MessageFormat = Request("AT+CMGF=1", , 2, vbCrLf)
    Case PDUFormat
        SMS_MessageFormat = Request("AT+CMGF=0", , 2, vbCrLf)
    Case ReadFormat
        mAnswer = Request("AT+CMGF?", , 2, vbCrLf)
        mAnswer = Trim$(MvField(mAnswer, 2, ":"))
        Select Case mAnswer
        Case 0
            SMS_MessageFormat = "PDU"
        Case 1
            SMS_MessageFormat = "TEXT"
        End Select
    End Select
    RaiseEvent Response(SMS_MessageFormat)
    Err.Clear
End Function
Public Function SMS_MemoryStorage(SelectMemory As SmsMemoryStorageEnum) As String
    On Error Resume Next
    ' select phonebook memory
    Dim mAnswer As String
    Select Case SelectMemory
    Case SimMemory
        mAnswer = Request("AT+CPMS=" & Me.Quote & "SM" & Me.Quote & "," & Me.Quote & "SM" & Me.Quote & "," & Me.Quote & "SM" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "SM"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "SM"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "SM"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case MobileEquipmentMemory
        mAnswer = Request("AT+CPMS=" & Me.Quote & "ME" & Me.Quote & "," & Me.Quote & "ME" & Me.Quote & "," & Me.Quote & "ME" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "ME"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "ME"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "ME"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case BothMemories
        mAnswer = Request("AT+CPMS=" & Me.Quote & "MT" & Me.Quote & "," & Me.Quote & "MT" & Me.Quote & "," & Me.Quote & "MT" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "MT"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "MT"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "MT"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case ReadMemorySetting
        mAnswer = Request("AT+CPMS?", "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mReadDelete = MvField(mAnswer, 1, ",")
        mReadDeleteUsed = Val(MvField(mAnswer, 2, ","))
        mReadDeleteCapacity = Val(MvField(mAnswer, 3, ","))
        mWriteSend = MvField(mAnswer, 4, ",")
        mWriteSendUsed = Val(MvField(mAnswer, 5, ","))
        mWriteSendCapacity = Val(MvField(mAnswer, 6, ","))
        mReceived = MvField(mAnswer, 7, ",")
        mReceivedUsed = Val(MvField(mAnswer, 8, ","))
        mReceivedCapacity = Val(MvField(mAnswer, 9, ","))
        mAnswer = "OK"
    Case BroadcastMessageStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "BM" & Me.Quote & "," & Me.Quote & "BM" & Me.Quote & "," & Me.Quote & "BM" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "BM"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "BM"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "BM"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case StatusReportStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "SR" & Me.Quote & "," & Me.Quote & "SR" & Me.Quote & "," & Me.Quote & "SR" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "SR"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "SR"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "SR"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case TerminalAdapterStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "TA" & Me.Quote & "," & Me.Quote & "TA" & Me.Quote & "," & Me.Quote & "TA" & Me.Quote, "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "TA"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "TA"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "TA"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    End Select
    SMS_MemoryStorage = mAnswer
    RaiseEvent Response(SMS_MemoryStorage)
    Err.Clear
End Function
Public Function SMS_CentreNumber(CentreNumberAction As CentreNumberEnum, Optional SMSC As String = "") As String
    On Error Resume Next
    ' centre number management
    Dim mAnswer As String
    Select Case CentreNumberAction
    Case SetCentreNumber
        mAnswer = Request("AT+CSCA=" & Me.Quote & SMSC & Me.Quote, "", 2, vbCrLf)
    Case ReadCentreNumber
        mAnswer = Request("AT+CSCA?", "", 2, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = MvField(mAnswer, 1, ",")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
    End Select
    SMS_CentreNumber = mAnswer
    RaiseEvent Response(SMS_CentreNumber)
    Err.Clear
End Function
Public Property Get SMS_ReadDeleteStorage() As String
    On Error Resume Next
    SMS_ReadDeleteStorage = mReadDelete
    Err.Clear
End Property
Public Property Get SMS_ReadDeleteStorageUsed() As Long
    On Error Resume Next
    SMS_ReadDeleteStorageUsed = mReadDeleteUsed
    Err.Clear
End Property
Public Property Get SMS_ReadDeleteStorageCapacity() As Long
    On Error Resume Next
    SMS_ReadDeleteStorageCapacity = mReadDeleteCapacity
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorage() As String
    On Error Resume Next
    SMS_WriteSendStorage = mWriteSend
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorageUsed() As Long
    On Error Resume Next
    SMS_WriteSendStorageUsed = mWriteSendUsed
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorageCapacity() As Long
    On Error Resume Next
    SMS_WriteSendStorageCapacity = mWriteSendCapacity
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorage() As String
    On Error Resume Next
    SMS_ReceivedStorage = mReceived
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorageUsed() As Long
    On Error Resume Next
    SMS_ReceivedStorageUsed = mReceivedUsed
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorageCapacity() As Long
    On Error Resume Next
    SMS_ReceivedStorageCapacity = mReceivedCapacity
    Err.Clear
End Property
Public Property Get PhoneBook_Used() As Long
    On Error Resume Next
    PhoneBook_Used = mpbUsed
    Err.Clear
End Property
Public Property Get PhoneBook_Capacity() As Long
    On Error Resume Next
    PhoneBook_Capacity = mpbCapacity
    Err.Clear
End Property
Public Property Get PhoneBook_Memory() As String
    On Error Resume Next
    ' update property
    PhoneBook_Memory = mpbMemory
    Err.Clear
End Property
Public Property Get SubscriberNumber() As String
    On Error Resume Next
    ' get the subscriber number
    SubscriberNumber = Request("AT+CNUM", , 2, vbCrLf)
    RaiseEvent Response(SubscriberNumber)
    Err.Clear
End Property
Public Property Get InternationalMobileSubscriberIdentity() As String
    On Error Resume Next
    ' get the international mobile subscriber identity
    InternationalMobileSubscriberIdentity = Request("AT+CIMI", , 2, vbCrLf)
    RaiseEvent Response(InternationalMobileSubscriberIdentity)
    Err.Clear
End Property
Public Function SMS_ReadMessageEntry(msgIndex As Long) As String
    On Error Resume Next
    ' read the message stored at location
    Dim mAnswer As String
    Dim tmpAns As String
    Dim msgType As String
    Dim msgFrom As String
    Dim msgDate As String
    Dim msgTime As String
    Dim msgContents As String
    Dim pPos As Long
    mAnswer = Request("AT+CMGR=" & msgIndex, , , vbCrLf)
    mAnswer = MvRest(mAnswer, 2, Me.VM)
    mAnswer = Trim$(Replace$(mAnswer, "+CMGR:", ""))
    tmpAns = MvField(mAnswer, 1, VM)
    Select Case tmpAns
    Case "OK"
        mAnswer = ""
    Case Else
        msgType = MvField(mAnswer, 1, ",")
        msgType = Replace$(msgType, Quote, "")
        msgFrom = MvField(mAnswer, 2, ",")
        msgFrom = Replace$(msgFrom, Quote, "")
        msgDate = MvField(mAnswer, 4, ",")
        msgDate = Replace$(msgDate, Quote, "")
        msgDate = FixSmsDate(msgDate)
        msgTime = MvField(mAnswer, 5, ",")
        msgTime = MvField(msgTime, 1, VM)
        msgTime = Replace$(msgTime, Quote, "")
        pPos = InStr(1, msgTime, "+")
        If pPos > 0 Then
            msgTime = Left$(msgTime, pPos - 1)
        End If
        msgContents = MvRest(mAnswer, 2, VM)
        msgContents = Replace$(msgContents, VM & VM & "OK" & VM, "")
        msgContents = RemAllVM(msgContents)
        mAnswer = msgIndex & FM & msgType & FM & msgFrom & FM & msgDate & " " & msgTime & FM & msgContents
    End Select
    SMS_ReadMessageEntry = mAnswer
    RaiseEvent Response(SMS_ReadMessageEntry)
    Err.Clear
End Function
Public Sub SMS_ListView(progBar As ProgressBar, LstView As ListView, Optional mIcon As String = "", Optional mSmallIcon As String = "", Optional msgType As SmsTypesEnum)
    On Error Resume Next
    ' load contents of the selected message store messages
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    Dim lstItem As ListItem
    Dim spLine() As String
    Dim nCollection As New Collection
    Set nCollection = SMS_ReadMessages
    progBar.Max = nCollection.Count
    progBar.Min = 0
    progBar.Value = 0
    ' create headings
    LstViewMakeHeadings LstView, "Msg ID,Cellphone No.,Contents,Time"
    ' loop through the messages starting from location 1 to the full capacity of the phone
    For rsCnt = 1 To nCollection.Count
        progBar.Value = rsCnt
        ' read entry at specified index
        phEntry = nCollection(rsCnt)
        ' if successfull return msg index, type,cellnumber,date,message
        If Len(phEntry) > 0 Then
            spLine = Split(phEntry, FM)
            Select Case msgType
            Case Rec
                If spLine(1) = "REC READ" Or spLine(1) = "REC UNREAD" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case RecRead
                If spLine(1) = "REC READ" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case RecUnread
                If spLine(1) = "REC UNREAD" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case StoSent
                If spLine(1) = "STO SENT" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case StoUnsent
                If spLine(1) = "STO UNSENT" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            End Select
AddLine:
            Set lstItem = LstView.ListItems.Add(, , spLine(0))
            lstItem.SubItems(1) = spLine(2)
            lstItem.SubItems(2) = spLine(4)
            lstItem.SubItems(3) = spLine(3)
            If Len(mIcon) > 0 Then lstItem.Icon = mIcon
            If Len(mSmallIcon) > 0 Then lstItem.SmallIcon = mSmallIcon
        End If
NextLine:
        DoEvents
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Function MvRest(ByVal strData As String, Optional ByVal startPos As Long = 1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    ' get the string from a substring position to the end of the
    ' delimited string
    Dim spData() As String
    Dim spCnt As Long
    Dim intLoop As Long
    Dim strL As String
    Dim strM As String
    MvRest = ""
    strM = ""
    If Len(Delim) = 0 Then Delim = Me.VM
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    spData = Split(strData, Delim)
    spCnt = UBound(spData)
    Select Case startPos
    Case -1
        MvRest = Trim$(spData(spCnt))
    Case Else
        strL = ""
        startPos = startPos - 1
        For intLoop = startPos To spCnt
            strL = spData(intLoop)
            If intLoop = spCnt Then
                strM = strM & strL
            Else
                strM = strM & strL & Delim
            End If
        Next
        MvRest = strM
    End Select
    Err.Clear
End Function
Public Function SMS_DeleteEntry(Location As Long) As String
    On Error Resume Next
    ' delete a sms and refresh if ok
    Dim mAnswer As String
    mAnswer = Request("AT+CMGD=" & Location, , 2, vbCrLf)
    SMS_DeleteEntry = mAnswer
    If mAnswer = "OK" Then Call SMS_MemoryStorage(ReadMemorySetting)
    RaiseEvent Response(SMS_DeleteEntry)
    Err.Clear
End Function
Public Function SMS_ReadMessages() As Collection
    On Error Resume Next
    ' read all sms messages from selected memory
    Dim mAnswer As String
    Dim nCollection As New Collection
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsLines() As String
    Dim msgIndex As String
    Dim msgType As String
    Dim msgFrom As String
    Dim msgDate As String
    Dim msgTime As String
    Dim msgContents As String
    Dim pPos As Long
    mAnswer = Me.SMS_MessageFormat(TextFormat)
    If mAnswer = "OK" Then
        mAnswer = Me.Request("AT+CMGL=" & Quote & "ALL" & Quote, , , vbCrLf)
    End If
    rsLines = Split(mAnswer, "+CMGL:")
    rsTot = UBound(rsLines)
    For rsCnt = 0 To rsTot
        rsLines(rsCnt) = Trim$(rsLines(rsCnt))
        rsStr = rsLines(rsCnt)
        msgIndex = MvField(rsStr, 1, ",")
        If IsNumeric(msgIndex) = True Then
            msgType = MvField(rsStr, 2, ",")
            msgType = Replace$(msgType, Quote, "")
            msgFrom = MvField(rsStr, 3, ",")
            msgFrom = Replace$(msgFrom, Quote, "")
            msgDate = MvField(rsStr, 5, ",")
            msgDate = Replace$(msgDate, Quote, "")
            msgDate = FixSmsDate(msgDate)
            msgTime = MvField(rsStr, 6, ",")
            msgTime = MvField(msgTime, 1, VM)
            msgTime = Replace$(msgTime, Quote, "")
            pPos = InStr(1, msgTime, "+")
            If pPos > 0 Then
                msgTime = Left$(msgTime, pPos - 1)
            End If
            msgContents = MvRest(rsStr, 2, VM)
            msgContents = Replace$(msgContents, VM & VM & "OK" & VM, "")
            msgContents = RemAllVM(msgContents)
            rsStr = msgIndex & FM & msgType & FM & msgFrom & FM & msgDate & " " & msgTime & FM & msgContents
            nCollection.Add rsStr
        End If
    Next
    Set SMS_ReadMessages = nCollection
    Err.Clear
End Function
Private Function SMS_SendSmall(sNumber As String, sMessage As String) As String
    On Error Resume Next
    ' send an sms to a phone, set textmode format just in case
    Dim mAnswer As String
    mAnswer = Request("AT+CMGS=" & Quote & sNumber & Quote, , 2, vbCr)
    Select Case mAnswer
    Case ">"
        mAnswer = Request(sMessage, , 4, Chr$(26))
    Case Else
        mAnswer = "ERROR"
    End Select
    SMS_SendSmall = mAnswer
    RaiseEvent Response(SMS_SendSmall)
    Err.Clear
End Function
Public Function SMS_Send(sNumber As String, sMessage As String) As String
    On Error Resume Next
    ' send an sms to a phone, set textmode format just in case
    Dim mAnswer As String
    Dim msgSent As Long
    Dim msgOne As Long
    Dim msgTwo As Long
    msgOne = 0
    msgTwo = 0
    mAnswer = SMS_MessageFormat(TextFormat)
    If mAnswer = "OK" Then
        If Len(sMessage) > 160 Then
            mAnswer = SMS_SendSmall(sNumber, Mid$(sMessage, 1, 160))
            Select Case LCase$(mAnswer)
            Case "ok"
                msgOne = 0
            Case Else
                msgOne = 1
            End Select
            mAnswer = SMS_SendSmall(sNumber, Mid$(sMessage, 161, 160))
            Select Case LCase$(mAnswer)
            Case "ok"
                msgTwo = 0
            Case Else
                msgTwo = 1
            End Select
            msgSent = msgOne + msgTwo
            If msgSent = 0 Then
                mAnswer = "OK"
            Else
                mAnswer = "ERROR"
            End If
        Else
            mAnswer = SMS_SendSmall(sNumber, sMessage)
        End If
    Else
        mAnswer = "ERROR"
    End If
    SMS_Send = mAnswer
    RaiseEvent Response(SMS_Send)
    Err.Clear
End Function
Private Function RemAllVM(ByVal StrString As String) As String
    On Error Resume Next
    Dim strSize As Long
    Dim strLast As String
    Dim tmpstring As String
    tmpstring = StrString
    strLast = Right$(tmpstring, 1)
    Do While strLast = VM
        strSize = Len(tmpstring) - 1
        tmpstring = Left$(tmpstring, strSize)
        strLast = Right$(tmpstring, 1)
    Loop
    RemAllVM = tmpstring
    Err.Clear
End Function
Public Function PhoneBook_Export(progBar As ProgressBar, exportFile As String) As Boolean
    On Error Resume Next
    ' load contents of the selected phonebook
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    'Dim lstItem As ListItem
    'Dim spLine() As String
    'Dim lstTotal As Long
    Dim pUsed As Long
    Dim pCapacity As Long
    Dim pWritten As Long
    pWritten = 0
    If FileExists(exportFile) = True Then Kill exportFile
    phEntry = PhoneBook_MemoryStorage(ReadPhoneBookSetting)
    If phEntry = "ERROR" Then
        PhoneBook_Export = False
    Else
        pCapacity = PhoneBook_Capacity
        pUsed = PhoneBook_Used
        progBar.Max = pCapacity
        progBar.Min = 0
        progBar.Value = 0
        FileUpdate exportFile, "Index,Cellphone No,Full Name", "a"
        For rsCnt = 1 To pCapacity
            progBar.Value = rsCnt
            ' read entry at specified index
            phEntry = Me.PhoneBook_ReadEntry(rsCnt)
            ' if successfull return index,cellno,fullname
            If Len(phEntry) > 0 Then
                pWritten = pWritten + 1
                FileUpdate exportFile, phEntry, "a"
            End If
            ' ensure that if we have reached the used limit, exit the loop
            ' we do not want to read the empty contacts anymore
            If pWritten = pUsed Then Exit For
            DoEvents
        Next
    End If
    progBar.Value = 0
    PhoneBook_Export = True
    Err.Clear
End Function
Private Function FileExists(ByVal Filename As String) As Boolean
    On Error Resume Next
    FileExists = False
    If Len(Filename) = 0 Then
        Err.Clear
        Exit Function
    End If
    FileExists = IIf(Dir$(Filename) <> "", True, False)
    Err.Clear
End Function
Public Function PhoneBook_Import(progBar As ProgressBar, importFile As String) As Long
    On Error Resume Next
    Dim phEntry As String
    Dim pWritten As Long
    Dim fData As String
    Dim fLines() As String
    Dim fTot As Long
    Dim fCnt As Long
    Dim cName As String
    Dim cNumber As String
    Dim cIndex As String
    Dim cLocation As String
    Dim cResult As String
    fData = FileData(importFile)
    fLines = Split(fData, vbNewLine)
    fTot = UBound(fLines)
    progBar.Max = fTot + 1
    progBar.Min = 0
    progBar.Value = 0
    pWritten = 0
    For fCnt = 0 To fTot + 1
        progBar.Value = fCnt + 1
        phEntry = fLines(fCnt)
        If Len(phEntry) = 0 Then GoTo NextRow
        cIndex = MvField(phEntry, 1, ",")
        cNumber = MvField(phEntry, 2, ",")
        cName = MvField(phEntry, 3, ",")
        If cIndex = "Index" Then GoTo NextRow
        cLocation = Me.PhoneBook_EntryExists(cName, cNumber)
        If cLocation = "0" Then
            cResult = PhoneBook_AddEntry(cNumber, cName)
            If cResult = "OK" Then
                pWritten = pWritten + 1
            End If
        End If
NextRow:
        DoEvents
    Next
    PhoneBook_Import = pWritten
    Err.Clear
End Function
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
