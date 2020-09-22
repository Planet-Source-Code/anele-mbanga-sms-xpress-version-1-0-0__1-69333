VERSION 5.00
Begin VB.Form frmNewSMS 
   Caption         =   "New SMS"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewSMS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReply 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   320
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   10935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      ToolTipText     =   "Send the message"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveDraft 
      Caption         =   "Draft"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      ToolTipText     =   "Save message to drafts"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOutbox 
      Caption         =   "Outbox"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      ToolTipText     =   "Save message to outbox, such will be sent after 30 seconds"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTo 
      Caption         =   "Contacts"
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      ToolTipText     =   "Select recipients from contacts"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtTotal 
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "320"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      ToolTipText     =   "Add a recipient to sms to"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   1575
      Left            =   1200
      MaxLength       =   320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   10935
   End
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reply To"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient(s)"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNewSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
    On Error Resume Next
    ' specify a number to send the message to
    Dim sNumber As String
    sNumber = InputBox("Please enter the number(s) that you want to send a sms to separated by a comma.", "Add Numbers")
    If Len(sNumber) = 0 Then Exit Sub
    If Len(txtTo.Text) = 0 Then
        txtTo.Text = sNumber
    Else
        txtTo.Text = txtTo.Text & "," & sNumber
    End If
    Err.Clear
End Sub
Private Sub cmdOutbox_Click()
    On Error Resume Next
    Call frmSMSExchange.dbExchange_Messages(-1, txtTo.Text, txtMsg.Text, Now(), "OutBox", AddGroup)
    Unload Me
    Err.Clear
End Sub
Private Sub cmdSaveDraft_Click()
    On Error Resume Next
    ' save the message to the outbox, this will be sent a minute from current time
    Call frmSMSExchange.dbExchange_Messages(-1, txtTo.Text, txtMsg.Text, Now(), "Draft", AddGroup)
    Unload Me
    Err.Clear
End Sub
Private Sub cmdSend_Click()
    On Error Resume Next
    ' send sms to receipients
    Dim rsCnt As Long
    Dim rsTot As Long
    If IsBlank(txtTo, "recipients") = True Then Exit Sub
    If IsBlank(txtMsg, "message") = True Then Exit Sub
    Dim spNumbers() As String
    spNumbers = Split(txtTo.Text, ",")
    rsTot = UBound(spNumbers)
    ProgBarInit frmSMSExchange.progBar, rsTot + 1
    For rsCnt = 0 To rsTot
        frmSMSExchange.progBar.Value = rsCnt + 1
        StatusMessage frmSMSExchange, "Sending message to " & spNumbers(rsCnt)
        frmSMSExchange.SendSMSOcx spNumbers(rsCnt), txtMsg.Text
        DoEvents
    Next
    frmSMSExchange.progBar.Value = 0
    StatusMessage frmSMSExchange
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    ' the form should stay the same, when maximized or resized
    If frmNewSMS.WindowState = 2 Then frmNewSMS.WindowState = 0
    Err.Clear
End Sub
Private Sub txtMsg_Change()
    On Error Resume Next
    ' display how many characters are left and enable suitable buttons
    If Len(txtMsg.Text) = 0 Then
        cmdSend.Enabled = False
        cmdSaveDraft.Enabled = False
        cmdOutbox.Enabled = False
    Else
        cmdSend.Enabled = True
        cmdSaveDraft.Enabled = True
        cmdOutbox.Enabled = True
    End If
    txtTotal.Text = 320 - Len(txtMsg.Text)
    Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FormMoveToNextControl KeyAscii
    Err.Clear
End Sub
