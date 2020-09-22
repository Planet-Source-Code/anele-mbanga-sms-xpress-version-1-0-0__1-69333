VERSION 5.00
Begin VB.Form frmContacts 
   Caption         =   "Add Phonebook Contact"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContacts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Add a new contact"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      ToolTipText     =   "Done with adding contacts"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Save contact"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtFullName 
      Height          =   315
      Left            =   1320
      MaxLength       =   22
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox txtNumber 
      Height          =   315
      Left            =   1320
      MaxLength       =   16
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cellphone No"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdApply_Click()
    On Error Resume Next
    If IsBlank(txtFullName, "full name") = True Then Exit Sub
    If IsBlank(txtNumber, "cellphone number") = True Then Exit Sub
    If InStr(1, txtFullName.Text, ",") > 0 Then
        Call MyPrompt("The name of the contact cannot have a comma.", "o", "e", "Contact Error")
        Err.Clear
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frmSMSExchange.CheckSignal.Enabled = False
    frmSMSExchange.Scheduler.Enabled = False
    txtNumber.Text = ExtractNumbers(txtNumber.Text)
    Dim mResult As String
    StatusMessage frmSMSExchange, "Searching for the contact, please wait..."
    mResult = frmSMSExchange.GSM.PhoneBook_EntryExists(txtFullName.Text, txtNumber.Text)
    Select Case mResult
    Case "0"
RetryAdd:
        StatusMessage frmSMSExchange, "Adding contact to phonebook, please wait..."
        mResult = frmSMSExchange.GSM.PhoneBook_AddEntry(txtNumber.Text, txtFullName)
    Case Else
        Call MyPrompt("This contact already exists in the selected memory." & vbCr & vbCr & _
        "Name: " & txtFullName.Text & vbCr & _
        "Cell No: " & txtNumber.Text & vbCr & _
        "Location: " & mResult, "o", "e", "Phonebook Entry Exists")
        Screen.MousePointer = vbDefault
        frmSMSExchange.CheckSignal.Enabled = True
        frmSMSExchange.Scheduler.Enabled = True
        StatusMessage frmSMSExchange
        Err.Clear
        Exit Sub
    End Select
    If mResult = "OK" Then
        txtFullName.Text = ""
        txtNumber.Text = ""
        Screen.MousePointer = vbDefault
        frmSMSExchange.CheckSignal.Enabled = True
        frmSMSExchange.Scheduler.Enabled = True
        StatusMessage frmSMSExchange
        Err.Clear
        Exit Sub
    Else
        resp = MyPrompt("This contact could not be written to the phonebook, " & LCase$(mResult), "rc", "e", "Phonebook Error")
        If resp = vbCancel Then
            Screen.MousePointer = vbDefault
            frmSMSExchange.CheckSignal.Enabled = True
            frmSMSExchange.Scheduler.Enabled = True
            StatusMessage frmSMSExchange
            Err.Clear
            Exit Sub
        End If
        GoTo RetryAdd
    End If
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdNew_Click()
    On Error Resume Next
    txtFullName.Text = ""
    txtNumber.Text = ""
    txtFullName.SetFocus
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    ' the form should stay the same, when maximized or resized
    If Me.WindowState = 2 Then Me.WindowState = 0
    Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FormMoveToNextControl KeyAscii
    Err.Clear
End Sub
Private Sub txtFullName_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtFullName, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtNumber, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtFullName_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtFullName
    Err.Clear
End Sub
Private Sub txtNumber_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtNumber
    Err.Clear
End Sub
Private Sub txtFullName_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtFullName.Text) = 0 Then Exit Sub
    txtFullName.Text = ProperCase(txtFullName.Text)
    Err.Clear
End Sub
Private Sub txtNumber_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtNumber.Text) = 0 Then Exit Sub
    txtNumber.Text = ProperCase(txtNumber.Text)
    Err.Clear
End Sub
