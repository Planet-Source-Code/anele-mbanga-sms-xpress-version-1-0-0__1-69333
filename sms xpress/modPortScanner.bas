Attribute VB_Name = "modPortScanner"
Option Explicit
' this is used to scan all ports that are available and list them
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Ports(0 To 100) As PORT_INFO_2
'// API Structures
Private Type PORT_INFO_2
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type
Private Type API_PORT_INFO_2
    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long
End Type
'// These declarations are used to detect what Com ports
'// are available
'// API Declarations
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'// API Structures
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
'// API constants
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
'// This detects if a COM ports is available.
'// Used by the ListComPorts() procedure
'// Returns TRUE if the COM exists, FALSE if the COM does not exist
Public Function COMAvailable(COMNum As Long) As Boolean
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES
    'try to open the COM port
    'hCOM = CreateFile("COM" & COMNum & "", 0, FILE_SHARE_READ + FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    hCOM = CreateFile("\.\COM" & COMNum & "", 0, FILE_SHARE_READ + FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        COMAvailable = False
    Else
        COMAvailable = True
        'close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function
'// This procedure returns all the available ports
'// Used by the cmdGetAllPorts_Click() procedure
'// Use ServerName to specify the name of a Remote Workstation i.e. "//WIN95WKST"
'// or leave it blank "" to get the ports of the local Machine
Public Function GetAvailablePorts(ServerName As String) As Long
    Dim ret As Long
    Dim PortsStruct(0 To 100) As API_PORT_INFO_2
    Dim pcbNeeded As Long
    Dim pcReturned As Long
    Dim TempBuff As Long
    Dim i As Integer
    '// Get the amount of bytes needed to contain the data returned by the API call
    ret = EnumPorts(ServerName, 2, TempBuff, 0, pcbNeeded, pcReturned)
    '// Allocate the Buffer
    TempBuff = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    ret = EnumPorts(ServerName, 2, TempBuff, pcbNeeded, pcbNeeded, pcReturned)
    If ret Then
        '// Convert the returned String Pointer Values to VB String Type
        CopyMem PortsStruct(0), ByVal TempBuff, pcbNeeded
        For i = 0 To pcReturned - 1
            Ports(i).pDescription = LPSTRtoSTRING(PortsStruct(i).pDescription)
            Ports(i).pPortName = LPSTRtoSTRING(PortsStruct(i).pPortName)
            Ports(i).pMonitorName = LPSTRtoSTRING(PortsStruct(i).pMonitorName)
            Ports(i).fPortType = PortsStruct(i).fPortType
        Next
    End If
    GetAvailablePorts = pcReturned
    '// Free the Heap Space allocated for the Buffer
    If TempBuff Then HeapFree GetProcessHeap(), 0, TempBuff
End Function
Public Function TrimStr(strName As String) As String
    '// Finds a null then trims the string
    Dim x As Integer
    x = InStr(strName, vbNullChar)
    If x > 0 Then TrimStr = Left$(strName, x - 1) Else TrimStr = strName
End Function
Public Function LPSTRtoSTRING(ByVal lngPointer As Long) As String
    Dim lngLength As Long
    '// Get number of characters in string
    lngLength = lstrlenW(lngPointer) * 2
    '// Initialize string so we have something to copy the string into
    LPSTRtoSTRING = String$(lngLength, 0)
    '// Copy the string
    CopyMem ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength
    '// Convert to Unicode
    LPSTRtoSTRING = TrimStr(StrConv(LPSTRtoSTRING, vbUnicode))
End Function
