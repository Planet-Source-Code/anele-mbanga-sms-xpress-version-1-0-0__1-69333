VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMSExchange 
   Caption         =   "SMS Xpress"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMSExchange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer CheckSignal 
      Interval        =   60000
      Left            =   4440
      Top             =   960
   End
   Begin SmsXPress.GSM GSM 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   1508
      _ExtentY        =   1508
   End
   Begin VB.Timer Scheduler 
      Interval        =   1
      Left            =   4440
      Top             =   480
   End
   Begin VB.ComboBox cboGroups 
      Height          =   315
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin SmsXPress.isExplorerBar xpSMS 
      Align           =   3  'Align Left
      Height          =   9885
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   17436
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontCharset     =   0
      UxThemeText     =   0   'False
      Begin VB.PictureBox picSettings 
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3795
         ScaleWidth      =   3795
         TabIndex        =   11
         Top             =   5280
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtIMEI 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   2400
            Width           =   2175
         End
         Begin MSComctlLib.ProgressBar progSignal 
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            Top             =   2760
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.TextBox txtMaxSpeed 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtSettings 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtPort 
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            ToolTipText     =   "Set listed modem as default modem"
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox txtMCN 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Text            =   "+27831000002"
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IMEI"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   330
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Signal"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Settings"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   585
         End
         Begin VB.Label lblModem 
            BackStyle       =   0  'Transparent
            Caption         =   "Modem Name"
            Height          =   795
            Left            =   1320
            TabIndex        =   16
            Top             =   120
            Width           =   2160
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modem Name"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   960
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comm Port"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Centre Number"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max Speed"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   795
         End
      End
      Begin VB.PictureBox picPhone 
         Height          =   11055
         Left            =   840
         ScaleHeight     =   10995
         ScaleWidth      =   3795
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
         Begin MSComctlLib.TreeView treePhone 
            Height          =   5655
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   9975
            _Version        =   393217
            Indentation     =   882
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imgIcons"
            Appearance      =   1
         End
      End
   End
   Begin VB.TextBox tbListViewEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar progBar 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   9885
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstReport 
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imgIcons"
      SmallIcons      =   "imgIcons"
      ColHdrIcons     =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   8880
      Top             =   3120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   113
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":0442
            Key             =   "employee"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":0D1C
            Key             =   "report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1036
            Key             =   "account"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1910
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1B34
            Key             =   "new"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1F60
            Key             =   "draft"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":283A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":2B54
            Key             =   "exchange"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":342E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":3780
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":3AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":3E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":4276
            Key             =   "recycled"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":46C8
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":4B1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":4E34
            Key             =   "owed"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":570E
            Key             =   "supplier"
            Object.Tag             =   "supplier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":5FE8
            Key             =   "group"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":68C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":719C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":7A76
            Key             =   "open"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":8350
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":866A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":8ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":8F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":97E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":9B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":9E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":A1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":A530
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":A882
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":ABD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":B026
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":B378
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":B4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":B924
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":BA36
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":BB48
            Key             =   "reply"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":BCA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":C0F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":C68E
            Key             =   "sms"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":CF68
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":D842
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":DB5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":E436
            Key             =   "safe"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":ED10
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":F02A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":F37C
            Key             =   "inbox1"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":F6CE
            Key             =   "sentto"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":FA20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":FE72
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":101C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":10516
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":10868
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":10CBA
            Key             =   "sendsms"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":11254
            Key             =   "chq"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":11F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":123B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":12802
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":12C54
            Key             =   "messages"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":130A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":13200
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":13652
            Key             =   "contacts"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":13AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":13EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":14348
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1479A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":14BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":14F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":15358
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":157AA
            Key             =   "sentbox"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":15BFC
            Key             =   "com1"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":15F16
            Key             =   "com2"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":16230
            Key             =   "com3"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1654A
            Key             =   "com4"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":16864
            Key             =   "com5"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":16B7E
            Key             =   "com6"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":16E98
            Key             =   "com7"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":171B2
            Key             =   "com8"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":174CC
            Key             =   "com9"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":177E6
            Key             =   "com10"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":17B00
            Key             =   "com11"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":17E1A
            Key             =   "com13"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":18134
            Key             =   "com14"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1844E
            Key             =   "com15"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":18768
            Key             =   "com16"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":18A82
            Key             =   "com17"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":18D9C
            Key             =   "com18"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":190B6
            Key             =   "com21"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":193D0
            Key             =   "com22"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":196EA
            Key             =   "com23"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":19A04
            Key             =   "com24"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":19D1E
            Key             =   "com25"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1A038
            Key             =   "com26"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1A352
            Key             =   "com27"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1A66C
            Key             =   "com28"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1A986
            Key             =   "com30"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1ACA0
            Key             =   "com19"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1AFBA
            Key             =   "modem"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1B40C
            Key             =   "drafts1"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1B69E
            Key             =   "outbox1"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1B930
            Key             =   "outbox"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1BD82
            Key             =   "memory"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1C1D4
            Key             =   "disk"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1C626
            Key             =   "inbox"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1CA78
            Key             =   "resend"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1CECA
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1D31C
            Key             =   "modems"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1D76E
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1DBC0
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSMSExchange.frx":1DCD2
            Key             =   "cut"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuForward 
         Caption         =   "Forward"
         Begin VB.Menu mnuMyGroups 
            Caption         =   "Groups"
            Begin VB.Menu mnuGroups 
               Caption         =   "Groups"
               Index           =   0
            End
         End
         Begin VB.Menu mnuMyContacs 
            Caption         =   "Contacts"
            Begin VB.Menu mnuContacts 
               Caption         =   "Contacts"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuReply 
         Caption         =   "Reply"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Recycle"
      End
   End
End
Attribute VB_Name = "frmSMSExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tHt As LVHITTESTINFO
Private lstItem As ListItem
Private sGroupName As String
Private sName As String
Private sCellPhone As String
Public Enum GroupsEnum
    AddGroup = 0
    DeleteGroup = 1
    RecycleGroup = 2
End Enum
Private rsCnt As Long
Private rsTot As Long
Private Sub CopyContacts(sKey As String)
    On Error Resume Next
    Dim mSelected As String
    Dim rsTot As Long
    Select Case sKey
    Case "phonecontacts_copysim", "contacts_copysim"
        ' copy from phone to sim
        mSelected = LstViewCheckedToMV(lstReport, 1, GSM.VM)
        rsTot = MvCount(mSelected, GSM.VM)
        If rsTot = 0 Then
            Call MyPrompt("Please check the contacts to copy to the sim first.", "o", "e", "Copy Contacts")
            Err.Clear
            Exit Sub
        Else
            resp = MyPrompt("Are you sure that you want to copy the checked contacts to the sim card?", "yn", "q", "Copy " & rsTot & " Contacts")
            If resp = vbNo Then Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.PhoneBook_Memory <> "SM" Then
            StatusMessage Me, "Selecting sim card phonebook memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(SimPhoneBook)
        End If
        StatusMessage Me, "Exporting contact(s), please wait..."
        ExportContactsToFile
        StatusMessage Me, "Importing contact(s), to sim card please wait..."
        rsTot = GSM.PhoneBook_Import(progBar, App.Path & "\copy contacts.csv")
        StatusMessage Me
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        Call MyPrompt("This serves to confirm that the number of contacts that were copied is " & rsTot, "o", "i", "Copy Contacts")
    Case "phonecontacts_copypc", "memorycontacts_copypc"
        ' copy from phone,sim to computer
        mSelected = LstViewCheckedToMV(lstReport, 1, GSM.VM)
        rsTot = MvCount(mSelected, GSM.VM)
        If rsTot = 0 Then
            Call MyPrompt("Please check the contacts to copy to the computer first.", "o", "e", "Copy Contacts")
            Err.Clear
            Exit Sub
        Else
            resp = MyPrompt("Are you sure that you want to copy the checked contacts to the computer?", "yn", "q", "Copy " & rsTot & " Contacts")
            If resp = vbNo Then Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Exporting contact(s), please wait..."
        ExportContactsToFile
        StatusMessage Me, "Importing contact(s), to sim computer please wait..."
        rsTot = ComputerBook_Import(App.Path & "\copy contacts.csv")
        StatusMessage Me
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        Call MyPrompt("This serves to confirm that the number of contacts that were copied is " & rsTot, "o", "i", "Copy Contacts")
    Case "memorycontacts_copyphone", "contacts_copyphone"
        ' copy from sim to phone
        mSelected = LstViewCheckedToMV(lstReport, 1, GSM.VM)
        rsTot = MvCount(mSelected, GSM.VM)
        If rsTot = 0 Then
            Call MyPrompt("Please check the contacts to copy to the phone first.", "o", "e", "Copy Contacts")
            Err.Clear
            Exit Sub
        Else
            resp = MyPrompt("Are you sure that you want to copy the checked contacts to the phone?", "yn", "q", "Copy " & rsTot & " Contacts")
            If resp = vbNo Then Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.PhoneBook_Memory <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment phonebook memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
        End If
        StatusMessage Me, "Exporting contact(s), please wait..."
        ExportContactsToFile
        StatusMessage Me, "Importing contact(s), to mobile equipment memory please wait..."
        rsTot = GSM.PhoneBook_Import(progBar, App.Path & "\copy contacts.csv")
        StatusMessage Me
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        Call MyPrompt("This serves to confirm that the number of contacts that were copied is " & rsTot, "o", "i", "Copy Contacts")
    End Select
    Err.Clear
End Sub
Private Sub ExportContacts(sKey As String)
    On Error Resume Next
    Dim exportFile As String
    Select Case sKey
    Case "contacts_export"
        StatusMessage Me, "Exporting computer contacts, please wait..."
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Call ComputerBook_Export
        StatusMessage Me
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
    Case "phonecontacts_export"
        If GSM.PhoneBook_Memory <> "ME" Then
            Screen.MousePointer = vbHourglass
            CheckSignal.Enabled = False
            Scheduler.Enabled = False
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
            StatusMessage Me
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
            Screen.MousePointer = vbDefault
        End If
DoContactsExport:
        exportFile = App.Path & "\phone contacts.csv"
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Exporting phone contacts, please wait..."
        If GSM.PhoneBook_Export(progBar, exportFile) = True Then
            Call MyPrompt("The phone contacts export file has been successfully created.", "o", "i", "Phone Contacts")
        Else
            resp = MyPrompt("The phone contacts export file could not be created.", "rc", "q", "Phone Contacts")
            If resp = vbCancel Then Exit Sub
            GoTo DoContactsExport
        End If
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        StatusMessage Me
    Case "memorycontacts_export"
        If GSM.PhoneBook_Memory <> "SM" Then
            CheckSignal.Enabled = False
            Scheduler.Enabled = False
            StatusMessage Me, "Selecting sim card memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(SimPhoneBook)
            StatusMessage Me
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
        End If
        exportFile = App.Path & "\sim contacts.csv"
DoContactsExport1:
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Exporting sim card contacts, please wait..."
        If GSM.PhoneBook_Export(progBar, exportFile) = True Then
            Call MyPrompt("The sim card contacts export file has been successfully created.", "o", "i", "Sim Card Contacts")
        Else
            resp = MyPrompt("The sim card contacts export file could not be created.", "rc", "q", "Sim Card Contacts")
            If resp = vbCancel Then Exit Sub
            GoTo DoContactsExport1
        End If
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        StatusMessage Me
        Screen.MousePointer = vbDefault
    End Select
    Err.Clear
End Sub
Private Sub ExportMessages(sKey As String)
    On Error Resume Next
    Select Case sKey
    Case "phonemsg_export"
    Case "casememorymsg_export"
    End Select
    Err.Clear
End Sub
Private Sub ForwardMessage(sKey As String)
    On Error Resume Next
    Dim lstSelected As ListItem
    Dim mResult As String
    Select Case sKey
    Case "phonemsg_forward"
        Select Case lstReport.Tag
        Case "phoneinbox", "phoneoutbox", "phonesentbox"
        Case Else
            Call MyPrompt("You can only forward messages from the phone inbox, outbox and sent items.", "o", "e", "Forward Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to forward first.", "o", "e", "Forward Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from phone, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        StatusMessage Me
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        With frmNewSMS
            .Caption = "Forward SMS"
            .txtMsg.Text = MvField(mResult, 5, GSM.FM)
            .txtTo.Text = ""
            .txtReply.Text = ""
            .cmdSend.Enabled = True
            .cmdSaveDraft.Enabled = True
            .cmdOutbox.Enabled = True
            .Show vbModal
        End With
    Case "memorymsg_forward"
        Select Case lstReport.Tag
        Case "memoryinbox", "memoryoutbox", "memorysentbox"
        Case Else
            Call MyPrompt("You can only forward messages from the sim card inbox, outbox and sent items.", "o", "e", "Forward Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to forward first.", "o", "e", "Forward Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from sim card, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        StatusMessage Me
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        With frmNewSMS
            .Caption = "Forward SMS"
            .txtMsg.Text = MvField(mResult, 5, GSM.FM)
            .txtTo.Text = ""
            .txtReply.Text = ""
            .cmdSend.Enabled = True
            .cmdSaveDraft.Enabled = True
            .cmdOutbox.Enabled = True
            .Show vbModal
        End With
    End Select
    Err.Clear
End Sub
Private Sub ReplyToMessage(sKey As String)
    On Error Resume Next
    Dim lstSelected As ListItem
    Dim mResult As String
    Dim mPhone As String
    Dim mMessage As String
    Select Case sKey
    Case "memorymsg_reply"
        Select Case lstReport.Tag
        Case "memoryinbox"
        Case Else
            Call MyPrompt("You can only reply to messages from the sim card inbox.", "o", "e", "Reply Error")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to reply to first.", "o", "e", "Reply Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from phone, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        mPhone = MvField(mResult, 3, GSM.FM)
        mMessage = MvField(mResult, 5, GSM.FM)
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        With frmNewSMS
            .Caption = "Reply To SMS"
            .txtReply.Text = MvField(mResult, 5, GSM.FM)
            .txtTo.Text = mPhone
            .txtMsg.Text = ""
            .cmdSend.Enabled = True
            .cmdSaveDraft.Enabled = True
            .cmdOutbox.Enabled = True
            .Show vbModal
        End With
    Case "phonemsg_reply"
        Select Case lstReport.Tag
        Case "phoneinbox"
        Case Else
            Call MyPrompt("You can only reply to messages from the phone inbox.", "o", "e", "Reply Error")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to reply to first.", "o", "e", "Reply Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from phone, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        mPhone = MvField(mResult, 3, GSM.FM)
        mMessage = MvField(mResult, 5, GSM.FM)
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        With frmNewSMS
            .Caption = "Reply To SMS"
            .txtReply.Text = MvField(mResult, 5, GSM.FM)
            .txtTo.Text = mPhone
            .txtMsg.Text = ""
            .cmdSend.Enabled = True
            .cmdSaveDraft.Enabled = True
            .cmdOutbox.Enabled = True
            .Show vbModal
        End With
    End Select
    Err.Clear
End Sub
Private Sub ResendMessage(sKey As String)
    On Error Resume Next
    Dim lstSelected As ListItem
    Dim mResult As String
    Dim mPhone As String
    Dim mMessage As String
    Select Case sKey
    Case "memorymsg_resend"
        Select Case lstReport.Tag
        Case "memoryoutbox", "memorysentbox"
        Case Else
            Call MyPrompt("You can only resend messages from the sim card outbox and sent items.", "o", "e", "Resend Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to resend first.", "o", "e", "Resend Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from phone, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        StatusMessage Me, "Resending the message, please wait..."
        mPhone = MvField(mResult, 3, GSM.FM)
        mMessage = MvField(mResult, 5, GSM.FM)
        SendSMSOcx mPhone, mMessage
        StatusMessage Me
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
    Case "phonemsg_resend"
        Select Case lstReport.Tag
        Case "phoneoutbox", "phonesentbox"
        Case Else
            Call MyPrompt("You can only resend messages from the phone outbox and sent items.", "o", "e", "Resend Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to resend first.", "o", "e", "Resend Error")
            Err.Clear
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Reading the message from phone, please wait..."
        mResult = GSM.SMS_ReadMessageEntry(lstSelected.Text)
        StatusMessage Me, "Resending the message, please wait..."
        mPhone = MvField(mResult, 3, GSM.FM)
        mMessage = MvField(mResult, 5, GSM.FM)
        SendSMSOcx mPhone, mMessage
        StatusMessage Me
        Screen.MousePointer = vbDefault
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
    End Select
    Err.Clear
End Sub
Private Sub CheckSignal_Timer()
    On Error Resume Next
    ' check the signal of the modem/phone and display it
    Dim mSignal As String
    CheckSignal.Enabled = False
    Scheduler.Enabled = False
    If GSM.PortOpen = True Then
        mSignal = GSM.SignalQualityMeasure
        progSignal.Value = Val(mSignal)
        StatusMessage Me, lblModem.Caption & " (" & mSignal & "% Signal)", 2
    End If
    CheckSignal.Enabled = True
    Scheduler.Enabled = True
    Err.Clear
End Sub
Private Sub cmdApply_Click()
    On Error Resume Next
    ' save settings of default gsm modem to registry
    SaveSetting App.Title, "account", "mcn", txtMCN.Text
    SaveSetting App.Title, "account", "modem", lblModem.Caption
    SaveSetting App.Title, "account", "port", txtPort.Text
    SaveSetting App.Title, "account", "settings", txtSettings.Text
    SaveSetting App.Title, "account", "speed", txtMaxSpeed.Text
    GSM.LogFile = App.Path & "\gsm.txt"
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim mNode As Node
    Dim mSignal As String
    CheckSignal.Enabled = False
    Scheduler.Enabled = False
    xpSMS.Width = 4500
    ' display current settings and hide some menus
    frmSYS_Splash.lblLicenseTo.Caption = "Retrieving saved settings..."
    mnuFile.Visible = False
    txtMCN.Text = GetSetting(App.Title, "account", "mcn", "+27831000002")
    lblModem.Caption = GetSetting(App.Title, "account", "modem", "")
    txtPort.Text = GetSetting(App.Title, "account", "port", "")
    txtSettings.Text = GetSetting(App.Title, "account", "settings", "19200,n,8,1")
    txtMaxSpeed.Text = GetSetting(App.Title, "account", "speed", "19200")
    GSM.LogFile = App.Path & "\gsm.txt"
    SetNavigator
    lstReport.View = lvwIcon
    ' add a status bar and a progress bar
    modSMSExchange.AddStatusBar StatusBar1, progBar
    frmSYS_Splash.lblLicenseTo.Caption = "Opening the database..."
    Set dbExchange = dao.OpenDatabase(App.Path & "\smsexchange.mdb")
    ' the first panel should show the database name
    StatusMessage Me, App.Path & "\smsexchange.mdb", 1
    StatusMessage Me, lblModem.Caption, 2
    frmSYS_Splash.lblLicenseTo.Caption = "Setting the navigator..."
    treePhone.Nodes.Clear
    Set mNode = treePhone.Nodes.Add(, , "pc", "Computer", "computer", "computer")
    Set mNode = treePhone.Nodes.Add("pc", tvwChild, "modems", "Modems", "modems", "modems")
    mNode.Expanded = True
    RefreshModems
    Set mNode = treePhone.Nodes.Add("pc", tvwChild, "group", "Groups", "group", "group")
    mNode.Expanded = True
    treePhone.Nodes.Add "group", tvwChild, "groups_new", "New", "new", "new"
    treePhone.Nodes.Add "group", tvwChild, "groups_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "group", tvwChild, "groups_sendmsg", "Send SMS", "sms", "sms"
    Set mNode = treePhone.Nodes.Add("pc", tvwChild, "contacts", "Contacts", "contacts", "contacts")
    mNode.Expanded = True
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_new", "New", "new", "new"
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_sendmsg", "Send SMS", "sms", "sms"
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_export", "Export", "excel", "excel"
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_copysim", "Computer > Sim", "copy", "copy"
    treePhone.Nodes.Add "contacts", tvwChild, "contacts_copyphone", "Computer > Phone", "copy", "copy"
    Set mNode = treePhone.Nodes.Add("pc", tvwChild, "messages", "Messages", "messages", "messages")
    mNode.Expanded = True
    treePhone.Nodes.Add "messages", tvwChild, "newmsg", "New", "new", "new"
    treePhone.Nodes.Add "messages", tvwChild, "msgtsk_forward", "Forward", "forward", "forward"
    treePhone.Nodes.Add "messages", tvwChild, "msgtsk_resend", "Re Send", "resend", "resend"
    treePhone.Nodes.Add "messages", tvwChild, "msgtsk_reply", "Reply", "reply", "reply"
    treePhone.Nodes.Add "messages", tvwChild, "msgtsk_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "messages", tvwChild, "draft", "Drafts", "draft", "draft"
    treePhone.Nodes.Add "messages", tvwChild, "inbox", "Inbox", "inbox", "inbox"
    treePhone.Nodes.Add "messages", tvwChild, "outbox", "Outbox", "outbox", "outbox"
    treePhone.Nodes.Add "messages", tvwChild, "sentbox", "Sentbox", "sentbox", "sentbox"
    treePhone.Nodes.Add "messages", tvwChild, "recycled", "Recycled", "recycled", "recycled"
    Set mNode = treePhone.Nodes.Add(, , "phone", "Phone", "modem", "modem")
    Set mNode = treePhone.Nodes.Add("phone", tvwChild, "phonecontacts", "Contacts", "contacts", "contacts")
    mNode.Expanded = True
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_new", "New", "new", "new"
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_sendmsg", "Send SMS", "sms", "sms"
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_export", "Export", "excel", "excel"
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_copysim", "Phone > Sim", "copy", "copy"
    treePhone.Nodes.Add "phonecontacts", tvwChild, "phonecontacts_copypc", "Phone > Computer", "copy", "copy"
    Set mNode = treePhone.Nodes.Add("phone", tvwChild, "phonemessages", "Messages", "messages", "messages")
    mNode.Expanded = True
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonemsg_forward", "Forward", "forward", "forward"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonemsg_resend", "Re Send", "resend", "resend"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonemsg_reply", "Reply", "reply", "reply"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonemsg_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonemsg_export", "Export", "excel", "excel"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phoneinbox", "Inbox", "inbox", "inbox"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phoneoutbox", "Outbox", "outbox", "outbox"
    treePhone.Nodes.Add "phonemessages", tvwChild, "phonesentbox", "Sentbox", "sentbox", "sentbox"
    Set mNode = treePhone.Nodes.Add(, , "memory", "Sim", "disk", "disk")
    Set mNode = treePhone.Nodes.Add("memory", tvwChild, "memorycontacts", "Contacts", "contacts", "contacts")
    mNode.Expanded = True
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_new", "New", "new", "new"
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_sendmsg", "Send SMS", "sms", "sms"
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_export", "Export", "excel", "excel"
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_copyphone", "Sim > Phone", "copy", "copy"
    treePhone.Nodes.Add "memorycontacts", tvwChild, "memorycontacts_copypc", "Sim > Computer", "copy", "copy"
    Set mNode = treePhone.Nodes.Add("memory", tvwChild, "memorymessages", "Messages", "messages", "messages")
    mNode.Expanded = True
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorymsg_forward", "Forward", "forward", "forward"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorymsg_resend", "Re Send", "resend", "resend"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorymsg_reply", "Reply", "reply", "reply"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorymsg_delete", "Delete", "delete", "delete"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorymsg_export", "Export", "excel", "excel"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memoryinbox", "Inbox", "inbox", "inbox"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memoryoutbox", "Outbox", "outbox", "outbox"
    treePhone.Nodes.Add "memorymessages", tvwChild, "memorysentbox", "Sentbox", "sentbox", "sentbox"
    frmSYS_Splash.lblLicenseTo.Caption = "Setting the message centre number..."
    frmSYS_Splash.lblLicenseTo.Caption = "Connecting to the modem..."
    If GSM.Connect(txtPort.Text, txtMaxSpeed.Text) = "OK" Then
        Call GSM.SMS_CentreNumber(txtMCN.Text)
        frmSYS_Splash.lblLicenseTo.Caption = "Checking the signal quality..."
        mSignal = GSM.SignalQualityMeasure
        progSignal.Value = Val(mSignal)
        StatusMessage Me, lblModem.Caption & " (" & mSignal & "% Signal)", 2
        frmSYS_Splash.lblLicenseTo.Caption = "Setting new message indicators on..."
        GSM.SMS_NewMessageIndicate False
        frmSYS_Splash.lblLicenseTo.Caption = "Reading imei number of modem..."
        txtIMEI.Text = GSM.ModemSerialNumber
        frmSYS_Splash.lblLicenseTo.Caption = "Setting sms mode to text format..."
        Call GSM.SMS_MessageFormat(TextFormat)
        frmSYS_Splash.lblLicenseTo.Caption = "Setting sms preferred storage area..."
        Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
        frmSYS_Splash.lblLicenseTo.Caption = "Setting / returning phonebook preferred storage area..."
        Call GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
        Call GSM.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
        StatusMessage Me, GSM.PhoneBook_Memory & " Phonebook, Capacity " & GSM.PhoneBook_Capacity & ", Used " & GSM.PhoneBook_Used, 3
    End If
    frmSYS_Splash.lblLicenseTo.Caption = ""
    CheckSignal.Enabled = True
    Scheduler.Enabled = True
    Err.Clear
End Sub
Private Sub GSM_Response(ByVal Result As String)
    On Error Resume Next
    'Debug.Print Result
    Err.Clear
End Sub
Private Sub lstReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    ' change the order of the clicked column, sort in asc or desc order
    LstViewSwapSort lstReport, ColumnHeader
    Err.Clear
End Sub
Private Sub lstReport_DblClick()
    On Error Resume Next
    Dim sCols As String
    Select Case lstReport.Tag
    Case "phonebook_phone"
        ' the contacts listed belong to the phone memory
        sCols = LstViewColNames(lstReport)
        Call ListView_ScaleEdit(lstReport, tHt, tbListViewEdit)
        Call ListView_BeforeEdit(lstReport, tHt, tbListViewEdit)
        SaveReg "column", "", App.Path, App.Title
        sCols = MvField(sCols, tHt.lSubItem + 1, ",")
        Select Case sCols
        Case "Cellphone No", "Full Name"
            SaveReg "column", sCols, "account", App.Title
        Case Else
            tbListViewEdit.Visible = False
        End Select
    Case "contacts"
        ' when a contact is double clicked, edit details
        sCols = LstViewColNames(lstReport)
        Call ListView_ScaleEdit(lstReport, tHt, tbListViewEdit)
        Call ListView_BeforeEdit(lstReport, tHt, tbListViewEdit)
        SaveReg "column", "", App.Path, App.Title
        sCols = MvField(sCols, tHt.lSubItem + 1, ",")
        Select Case sCols
        Case "Name", "Cell Phone", "Group"
            SaveReg "column", sCols, "account", App.Title
        Case Else
            tbListViewEdit.Visible = False
        End Select
    Case "groups"
        ' list all contacts meeting criteria
        Set lstItem = lstReport.SelectedItem
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        Dim sGroup As String
        Dim rsContacts As dao.Recordset
        Caption = "SMS Xpress: Contacts For Group " & lstItem.Text
        LstViewMakeHeadings lstReport, "Name,Cell Phone,Group"
        lstReport.ColumnHeaders(2).Alignment = lvwColumnRight
        lstReport.Checkboxes = False
        lstReport.Tag = "contacts"
        Set rsContacts = dbExchange.OpenRecordset("select * from contacts where groupname = '" & lstItem.Text & "' order by name;")
        rsContacts.MoveLast
        rsTot = rsContacts.RecordCount
        ProgBarInit progBar, rsTot
        rsContacts.MoveFirst
        For rsCnt = 1 To rsTot
            progBar.Value = rsCnt
            sGroup = rsContacts!GroupName & ""
            sName = rsContacts!Name & ""
            sCellPhone = rsContacts!cellphone & ""
            Set lstItem = lstReport.ListItems.Add(, , sName, "employee", "employee")
            lstItem.SubItems(1) = sCellPhone
            lstItem.SubItems(2) = sGroup
            DoEvents
            rsContacts.MoveNext
        Next
        progBar.Value = 0
        rsContacts.Close
        StatusMessage Me, lstReport.ListItems.Count & " contacts(s) listed."
    End Select
    Err.Clear
End Sub
Private Sub lstReport_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case 46
        ' a use has selected to delete the group/contact by pressing the delete button
        Dim lstSelected As ListItem
        Dim sGroupName As String
        Dim mResult As String
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then Exit Sub
        sGroupName = lstSelected.Text
        Select Case lstReport.Tag
        Case "groups"
            resp = MsgBox("Are you sure that you want to delete this group?" & vbCr & vbCr & sGroupName, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
            If resp = vbNo Then Exit Sub
            dbExchange_Groups sGroupName, DeleteGroup
            lstReport.ListItems.Remove lstSelected.Index
        Case "contacts"
            sGroupName = lstSelected.ListSubItems(2)
            resp = MsgBox("Are you sure that you want to delete this contact from the database?" & vbCr & vbCr & sGroupName, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
            If resp = vbNo Then Exit Sub
            dbExchange_Contacts sGroupName, lstSelected.SubItems(1), lstSelected.SubItems(3), DeleteGroup
            lstReport.ListItems.Remove lstSelected.Index
        Case "phonebook_phone"
            resp = MsgBox("Are you sure that you want to delete this contact from the phonebook?" & vbCr & vbCr & lstSelected.SubItems(2), vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
            If resp = vbNo Then Exit Sub
            Screen.MousePointer = vbHourglass
            StatusMessage Me, "Deleting contact from phonebook, please wait..."
            mResult = GSM.PhoneBook_DeleteEntry(Val(lstSelected.Text))
            Screen.MousePointer = vbDefault
            StatusMessage Me
            If mResult = "OK" Then
                lstReport.ListItems.Remove lstSelected.Index
            Else
                Call MyPrompt("The selected contact could not be deleted.", "o", "e", "Delete Error")
            End If
        Case "phoneinbox", "phoneoutbox", "phonedraft", "phonesentbox", "memoryinbox", "memoryoutbox", "memorydraft", "memorysentbox"
            resp = MsgBox("Are you sure that you want to delete this message from storage?" & vbCr & vbCr & "Message index " & lstSelected.Text, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
            If resp = vbNo Then Exit Sub
            Screen.MousePointer = vbHourglass
            CheckSignal.Enabled = False
            Scheduler.Enabled = False
            StatusMessage Me, "Deleting message from storage, please wait..."
            mResult = GSM.SMS_DeleteEntry(Val(lstSelected.Text))
            Screen.MousePointer = vbDefault
            StatusMessage Me
            If mResult = "OK" Then
                lstReport.ListItems.Remove lstSelected.Index
            Else
                Call MyPrompt("The selected message could not be deleted.", "o", "e", "Delete Error")
            End If
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
        End Select
    Case 45
        Select Case lstReport.Tag
        Case "groups"
            ' a use has selected to add a new group/contact by pressing the insert button
            ' add a new group
            sGroupName = InputBox("Please enter the new name of the group to add.", "New Group")
            If Len(sGroupName) = 0 Then Exit Sub
            sGroupName = ProperCase(sGroupName)
            ' the group has been added
            If dbExchange_Groups(sGroupName, AddGroup) = True Then
                ' show all the groups
                xpSMS_ItemClick "", "groups_list"
            End If
        Case "contacts"
            xpSMS_ItemClick "", "contacts_new"
        Case "phonebook_phone"
            frmContacts.txtFullName.Text = ""
            frmContacts.txtNumber.Text = ""
            frmContacts.Show vbModal
        End Select
    End Select
    Err.Clear
End Sub
Private Sub mnuContacts_Click(Index As Integer)
    On Error Resume Next
    ' send selected sms to a particular contact
    Set lstItem = lstReport.SelectedItem
    If TypeName(lstItem) = "Nothing" Then Exit Sub
    SendSms_ToSomeOne "Name", mnuContacts(Index).Caption, lstItem.Text
    Err.Clear
End Sub
Private Sub mnuDelete_Click()
    On Error Resume Next
    'delete the sms
    xpSMS_ItemClick "", "msgtsk_delete"
    Err.Clear
End Sub
Private Sub mnuGroups_Click(Index As Integer)
    On Error Resume Next
    ' send sms to a particular group
    Set lstItem = lstReport.SelectedItem
    If TypeName(lstItem) = "Nothing" Then Exit Sub
    SendSms_ToSomeOne "GroupName", mnuGroups(Index).Caption, lstItem.Text
    Err.Clear
End Sub
Private Sub mnuReply_Click()
    On Error Resume Next
    ' resend the sms
    xpSMS_ItemClick "", "msgtsk_reply"
    Err.Clear
End Sub
Private Sub mnuSend_Click()
    On Error Resume Next
    ' resend the sms
    xpSMS_ItemClick "", "msgtsk_resend"
    Err.Clear
End Sub
Private Sub picPhone_Resize()
    On Error Resume Next
    treePhone.Appearance = ccFlat
    treePhone.BorderStyle = ccNone
    treePhone.Top = picPhone.ScaleTop
    treePhone.Left = picPhone.ScaleLeft
    treePhone.Height = picPhone.ScaleHeight
    treePhone.Width = picPhone.ScaleWidth
    Err.Clear
End Sub
Private Sub Scheduler_Timer()
    On Error Resume Next
    ' check with messages are in the outbox at the current time and resend everything
    Scheduler.Enabled = False
    If TypeName(dbExchange) = "Nothing" Then GoTo ExitTimer
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sNow As String
    Dim rsMessages As dao.Recordset
    Dim sID As String
    sNow = Now()
    qrySql = "Select ID From Messages Where MsgType = 'OutBox' And Messages.ProcessTime = '" & sNow & "';"
    Set rsMessages = dbExchange.OpenRecordset(qrySql)
    rsMessages.MoveLast
    rsTot = rsMessages.RecordCount
    If Err.Number = 3021 Then GoTo ExitTimer
    ProgBarInit progBar, rsTot
    rsMessages.MoveFirst
    For rsCnt = 1 To rsTot
        progBar.Value = rsCnt
        sID = rsMessages!ID & ""
        dbExchange_SendSMS sID, False
        DoEvents
        rsMessages.MoveNext
    Next
ExitTimer:
    If TypeName(rsMessages) <> "Nothing" Then rsMessages.Close
    progBar.Value = 0
    LstViewAutoResize lstReport
    Scheduler.Enabled = True
    Err.Clear
End Sub
Private Sub tbListViewEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim sCols As String
    Dim upRec() As String
    Dim mResult As String
    ' after editing of an item, resave to database
    If KeyCode = vbKeyEscape Then
        tbListViewEdit.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        sCols = ReadReg("column", "account", App.Title)
        Select Case sCols
        Case "Name"
            ' name of contact saved in the database
            tbListViewEdit.Text = ProperCase(tbListViewEdit.Text)
            Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
            upRec = LstViewGetRow(lstReport, tHt.lItem + 1)
            Call LstViewUpdate(upRec, lstReport, CStr(tHt.lItem + 1))
            dbExchange_Contacts upRec(3), upRec(2), upRec(4), AddGroup
        Case "Cell Phone"
            ' cellphone number stored in the database
            tbListViewEdit.Text = ExtractNumbers(tbListViewEdit.Text)
            Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
            upRec = LstViewGetRow(lstReport, tHt.lItem + 1)
            Call LstViewUpdate(upRec, lstReport, CStr(tHt.lItem + 1))
            dbExchange_Contacts upRec(3), upRec(2), upRec(4), AddGroup
        Case "Group"
            ' group name stored in the database
            tbListViewEdit.Text = ProperCase(tbListViewEdit.Text)
            ' if the group does not exist, add it
            dbExchange_Groups tbListViewEdit.Text, AddGroup
            Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
            upRec = LstViewGetRow(lstReport, tHt.lItem + 1)
            Call LstViewUpdate(upRec, lstReport, CStr(tHt.lItem + 1))
            dbExchange_Contacts upRec(3), upRec(2), upRec(4), AddGroup
        Case "Cellphone No"
            Screen.MousePointer = vbHourglass
            CheckSignal.Enabled = False
            Scheduler.Enabled = False
            ' cellphone from gadget / sim card
            tbListViewEdit.Text = ExtractNumbers(tbListViewEdit.Text)
            Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
            upRec = LstViewGetRow(lstReport, tHt.lItem + 1)
            Call LstViewUpdate(upRec, lstReport, CStr(tHt.lItem + 1))
            ' update contact on database
            StatusMessage Me, "Updating phonebook contact, please wait..."
            mResult = GSM.PhoneBook_WriteEntry(Val(upRec(1)), upRec(2), upRec(3))
            StatusMessage Me
            Screen.MousePointer = Default
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
            If mResult = "OK" Then
            Else
                Call MyPrompt("This contact details could not be updated.", "o", "e", "Phonebook Error")
                Err.Clear
                Exit Sub
            End If
        Case "Full Name"
            ' full name from gadget / sim card
            Screen.MousePointer = vbHourglass
            CheckSignal.Enabled = False
            Scheduler.Enabled = False
            ' cellphone from gadget / sim card
            tbListViewEdit.Text = ProperCase(tbListViewEdit.Text)
            Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
            upRec = LstViewGetRow(lstReport, tHt.lItem + 1)
            Call LstViewUpdate(upRec, lstReport, CStr(tHt.lItem + 1))
            ' update contact on database
            StatusMessage Me, "Updating phonebook contact, please wait..."
            mResult = GSM.PhoneBook_WriteEntry(Val(upRec(1)), upRec(2), upRec(3))
            StatusMessage Me
            Screen.MousePointer = Default
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
            If mResult = "OK" Then
            Else
                Call MyPrompt("This contact details could not be updated.", "o", "e", "Phonebook Error")
                Err.Clear
                Exit Sub
            End If
        End Select
        ' after saving the information
        LstViewAutoResize lstReport
        lstReport.ListItems(tHt.lItem + 1).EnsureVisible
        ' find the next column to edit
        sCols = LstViewColNames(lstReport)
        Select Case MvField(sCols, tHt.lSubItem + 1, ",")
        Case "Cell Phone", "Group", "Cellphone No", "Full Name"
            tbListViewEdit.Visible = True
            lstReport.ListItems(tHt.lItem + 1).Selected = True
            Call ListView_ScaleEdit(lstReport, tHt, tbListViewEdit)
            Call ListView_BeforeEdit(lstReport, tHt, tbListViewEdit)
            lstReport.SetFocus
            tbListViewEdit.SetFocus
        End Select
    End If
    Err.Clear
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    ' add a progress bar when the form is activated
    PutProgressBarInStatusBar StatusBar1, progBar, 5
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    ' resize controls inside the form
    lstReport.Top = Me.ScaleTop
    lstReport.Left = Me.ScaleLeft + xpSMS.Width
    lstReport.Width = Me.ScaleWidth - xpSMS.Width
    lstReport.Height = Me.ScaleHeight - StatusBar1.Height
    ResizeStatusBar Me, StatusBar1, progBar
    Err.Clear
End Sub
Private Sub SetNavigator()
    On Error Resume Next
    ' load navigation tools
    With xpSMS
        .DisableUpdates True
        .ClearStructure
        .SetImageList imgIcons
        ' add the modem group
        .AddGroup "settings", "Default Modem", , Collapsed, picSettings
        .AddGroup "phone", "Navigator", , Expanded, picPhone
        .AddDetailsGroup "Details", "Management of sms messages"
        .DisableUpdates False
    End With
    Err.Clear
End Sub
Private Sub treePhone_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim sKey As String
    Dim sOther As String
    Dim pCapacity As Long
    Dim pUsed As Long
    'Dim lstSelected As ListItem
    'Dim mResult As String
    'Dim mPhone As String
    'Dim mMessage As String
    'Dim exportFile As String
    sKey = MvField(Node.Key, 1, ",")
    sOther = Rest(Node.Key, 2, ",")
    cmdApply.Enabled = False
    ExportMessages sKey
    ExportContacts sKey
    ReplyToMessage sKey
    ResendMessage sKey
    ForwardMessage sKey
    DeleteMessage sKey
    CopyContacts sKey
    Select Case LCase$(sKey)
    Case "phone"
        Screen.MousePointer = vbHourglass
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Caption = "SMS Xpress: Phone"
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.SMS_ReceivedStorage <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
            StatusMessage Me
        End If
        If GSM.PhoneBook_Memory <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
            StatusMessage Me
        End If
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "memory"
        Screen.MousePointer = vbHourglass
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Caption = "SMS Xpress: Sim Card"
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.SMS_ReceivedStorage <> "SM" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(SimMemory)
            StatusMessage Me
        End If
        If GSM.PhoneBook_Memory <> "SM" Then
            StatusMessage Me, "Selecting sim card memory, please wait..."
            Call GSM.PhoneBook_MemoryStorage(SimPhoneBook)
            StatusMessage Me
        End If
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "modems"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Caption = "SMS Xpress: Modems"
        RefreshModems
    Case "memorysentbox"
        Caption = "SMS Xpress: Sim Card Sentbox"
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "SM" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(SimMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "outbox", "outbox", StoSent
        StatusMessage Me
        lstReport.Tag = "memorysentbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "phonesentbox"
        Caption = "SMS Xpress: Phone Sentbox"
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "outbox", "outbox", StoSent
        StatusMessage Me
        lstReport.Tag = "phonesentbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "memoryoutbox"
        Caption = "SMS Xpress: Sim Card Outbox"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "SM" Then
            StatusMessage Me, "Selecting sim card memory, please wait..."
            Call GSM.SMS_MemoryStorage(SimMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "outbox", "outbox", StoUnsent
        StatusMessage Me
        lstReport.Tag = "memoryoutbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "phoneoutbox"
        Caption = "SMS Xpress: Phone Outbox"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "outbox", "outbox", StoUnsent
        StatusMessage Me
        lstReport.Tag = "phoneoutbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "phoneinbox"
        Caption = "SMS Xpress: Phone Inbox"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "inbox", "inbox", Rec
        StatusMessage Me
        lstReport.Tag = "phoneinbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "phonemessages"
SelectMemory:
        Caption = "SMS Xpress: Phone Messages"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.SMS_ReceivedStorage <> "ME" Then
            StatusMessage Me, "Selecting mobile equipment memory, please wait..."
            sKey = GSM.SMS_MemoryStorage(MobileEquipmentMemory)
            If sKey <> "OK" Then
                resp = MyPrompt("The mobile equipment memory could not be selected.", "rc", "q", "Prefered Memory")
                If resp = vbCancel Then
                    StatusMessage Me
                    CheckSignal.Enabled = True
                    Scheduler.Enabled = True
                    Err.Clear
                    Exit Sub
                End If
                GoTo SelectMemory
            End If
        End If
        pCapacity = GSM.SMS_ReceivedStorageCapacity
        pUsed = GSM.SMS_ReceivedStorageUsed
        xpSMS.SetDetailsText "Imei: " & txtIMEI.Text & vbNewLine & _
        "Capacity: " & pCapacity & vbNewLine & _
        "Used: " & pUsed
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        StatusMessage Me
    Case "memorymessages"
SelectMemory1:
        Caption = "SMS Xpress: Sim Messages"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.SMS_ReceivedStorage <> "SM" Then
            StatusMessage Me, "Selecting sim card memory, please wait..."
            sKey = GSM.SMS_MemoryStorage(SimMemory)
            If sKey <> "OK" Then
                resp = MyPrompt("The sim card memory could not be selected.", "rc", "q", "Prefered Memory")
                If resp = vbCancel Then
                    StatusMessage Me
                    CheckSignal.Enabled = True
                    Scheduler.Enabled = True
                    Err.Clear
                    Exit Sub
                End If
                GoTo SelectMemory1
            End If
        End If
        StatusMessage Me
        pCapacity = GSM.SMS_ReceivedStorageCapacity
        pUsed = GSM.SMS_ReceivedStorageUsed
        xpSMS.SetDetailsText "Imei: " & txtIMEI.Text & vbNewLine & _
        "Capacity: " & pCapacity & vbNewLine & _
        "Used: " & pUsed
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
    Case "memoryinbox"
        Caption = "SMS Xpress: Sim Card Inbox"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        If GSM.SMS_ReceivedStorage <> "SM" Then
            StatusMessage Me, "Selecting sim card memory, please wait..."
            Call GSM.SMS_MemoryStorage(SimMemory)
        End If
        ' read messages from phone
        StatusMessage Me, "Loading messages, please wait..."
        GSM.SMS_ListView progBar, lstReport, "inbox", "inbox", Rec
        StatusMessage Me
        lstReport.Tag = "memoryinbox"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "potsmodem", "potsmodemtoserialport"
        ' view modem properties
        Dim vItems As Variant
        Dim vTemp As Variant
        Caption = "SMS Xpress Modem Properties: " & ProperCase$(sOther)
        ' Clear the ListView
        LstViewMakeHeadings lstReport, "Property,Value"
        lstReport.ListItems.Clear
        lstReport.View = lvwReport
        lstReport.Checkboxes = False
        ' Populate the ListView with the device's properties
        For Each vTemp In GetProperties(sKey, sOther)
            vItems = Split(vTemp, "^")
            lstReport.ListItems.Add(, , CStr(vItems(0))).SubItems(1) = vItems(1)
        Next
        LstViewAutoResize lstReport
        lstReport.Tag = "modems"
        ReadModem
        cmdApply.Enabled = True
    Case "exchange", "messages"
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        Caption = "SMS Xpress"
    Case "group"
        xpSMS_ItemClick "", "groups_list"
    Case "groups_new", "groups_delete", "groups_sendmsg", "contacts_new", "contacts_delete", "contacts_sendmsg", _
        "newmsg", "draft", "outbox", "sentbox", "recycled", "msgtsk_forward", "msgtsk_resend", "msgtsk_delete", _
        "phonecontacts_new", "phonecontacts_delete", "phonecontacts_sendmsg", "inbox", "msgtsk_reply"
        xpSMS_ItemClick "", Node.Key
    Case "memorycontacts_new"
        xpSMS_ItemClick "", "phonecontacts_new"
    Case "memorycontacts_delete"
        xpSMS_ItemClick "", "phonecontacts_delete"
    Case "memorycontacts_sendmsg"
        xpSMS_ItemClick "", "phonecontacts_sendmsg"
    Case "contacts"
        xpSMS_ItemClick "", "contacts_list"
    Case "phonecontacts"
        Caption = "SMS Xpress: Phone Contacts"
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Selecting phonebook of mobile, please wait..."
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        If GSM.PhoneBook_Memory <> "ME" Then
            If GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook) = "OK" Then
            Else
                Call MyPrompt("The phone memory cannot be accessed.", "o", "e", "Phone Memory")
                StatusMessage Me
                Screen.MousePointer = vbDefault
                CheckSignal.Enabled = True
                Scheduler.Enabled = True
                Err.Clear
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Checking phone book capacity, please wait..."
        Call GSM.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
        StatusMessage Me, GSM.PhoneBook_Memory & " Phonebook, Capacity " & GSM.PhoneBook_Capacity & ", Used " & GSM.PhoneBook_Used, 3
        pCapacity = GSM.PhoneBook_Capacity
        pUsed = GSM.PhoneBook_Used
        xpSMS.SetDetailsText "Imei: " & txtIMEI.Text & vbNewLine & _
        "Capacity: " & pCapacity & vbNewLine & _
        "Used: " & pUsed
        If Val(pUsed) = 0 Then
            StatusMessage Me
            Screen.MousePointer = vbDefault
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
            Err.Clear
            Exit Sub
        Else
            ' read contacts from phone
            StatusMessage Me, "Loading phone book contacts, please wait..."
            GSM.PhoneBook_ListView progBar, lstReport, pUsed, pCapacity, "employee", "employee"
        End If
        StatusMessage Me
        lstReport.Tag = "phonebook_phone"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    Case "memorycontacts"
        Caption = "SMS Xpress: Sim Card Contacts"
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Selecting phonebook of sim, please wait..."
        lstReport.ListItems.Clear
        lstReport.View = lvwIcon
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        If GSM.PhoneBook_Memory <> "SM" Then
            If GSM.PhoneBook_MemoryStorage(SimPhoneBook) = "OK" Then
                CheckSignal.Enabled = True
                Scheduler.Enabled = True
                Screen.MousePointer = vbDefault
            Else
                Call MyPrompt("The sim card memory cannot be accessed.", "o", "e", "Sim Card Memory")
                StatusMessage Me
                Screen.MousePointer = vbDefault
                Err.Clear
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Checking sim card phone book capacity, please wait..."
        Call GSM.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
        StatusMessage Me, GSM.PhoneBook_Memory & " Phonebook, Capacity " & GSM.PhoneBook_Capacity & ", Used " & GSM.PhoneBook_Used, 3
        pCapacity = GSM.PhoneBook_Capacity
        pUsed = GSM.PhoneBook_Used
        xpSMS.SetDetailsText "Imei: " & txtIMEI.Text & vbNewLine & _
        "Capacity: " & pCapacity & vbNewLine & _
        "Used: " & pUsed
        If Val(pUsed) = 0 Then
            StatusMessage Me
            lstReport.ListItems.Clear
            lstReport.View = lvwIcon
            Screen.MousePointer = vbDefault
            CheckSignal.Enabled = True
            Scheduler.Enabled = True
            Err.Clear
            Exit Sub
        Else
            ' read contacts from phone
            StatusMessage Me, "Loading phone book contacts, please wait..."
            GSM.PhoneBook_ListView progBar, lstReport, pUsed, pCapacity, "employee", "employee"
        End If
        StatusMessage Me
        lstReport.Tag = "phonebook_phone"
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
    End Select
    Err.Clear
End Sub
Private Sub DeleteMessage(sKey As String)
    On Error Resume Next
    Dim lstSelected As ListItem
    Dim mResult As String
    Select Case sKey
    Case "memorymsg_delete"
        Select Case lstReport.Tag
        Case "memoryinbox", "memoryoutbox", "memorydraft", "memorysentbox"
        Case Else
            Call MyPrompt("You can only delete messages from the sim inbox, outbox, drafts and sent items.", "o", "e", "Delete Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to delete first.", "o", "e", "Delete Error")
            Err.Clear
            Exit Sub
        End If
        resp = MsgBox("Are you sure that you want to delete this message from storage?" & vbCr & vbCr & "Message index " & lstSelected.Text, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
        If resp = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        StatusMessage Me, "Deleting message from storage, please wait..."
        mResult = GSM.SMS_DeleteEntry(Val(lstSelected.Text))
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        StatusMessage Me
        If mResult = "OK" Then
            lstReport.ListItems.Remove lstSelected.Index
        Else
            Call MyPrompt("The selected message could not be deleted.", "o", "e", "Delete Error")
        End If
    Case "phonemsg_delete"
        Select Case lstReport.Tag
        Case "phoneinbox", "phoneoutbox", "phonedraft", "phonesentbox"
        Case Else
            Call MyPrompt("You can only delete messages from the phone inbox, outbox, drafts and sent items.", "o", "e", "Delete Message")
            Err.Clear
            Exit Sub
        End Select
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the message to delete first.", "o", "e", "Delete Error")
            Err.Clear
            Exit Sub
        End If
        resp = MsgBox("Are you sure that you want to delete this message from storage?" & vbCr & vbCr & "Message index " & lstSelected.Text, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
        If resp = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Deleting message from storage, please wait..."
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        mResult = GSM.SMS_DeleteEntry(Val(lstSelected.Text))
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        Screen.MousePointer = vbDefault
        StatusMessage Me
        If mResult = "OK" Then
            lstReport.ListItems.Remove lstSelected.Index
        Else
            Call MyPrompt("The selected message could not be deleted.", "o", "e", "Delete Error")
        End If
    End Select
    Err.Clear
End Sub
Private Sub xpSMS_GroupClick(sGroup As String, bExpanded As Boolean)
    On Error Resume Next
    ' when a group is clicked, display appropriate caption
    Select Case sGroup
    Case "modems", "settings", "groups", "contacts", "phone"
        Caption = "SMS Xpress: " & ProperCase(sGroup)
        lstReport.View = lvwIcon
        lstReport.ListItems.Clear
    Case "msgs"
    End Select
    Err.Clear
End Sub
Private Sub xpSMS_GroupHover(sGroup As String)
    On Error Resume Next
    ' when hovering on a group change the details
    Select Case sGroup
    Case "modems"
        xpSMS.SetDetailsText "List of current available modems connected to the computer."
    Case "settings"
        xpSMS.SetDetailsText "The default modem that will be used to send smses etc."
    Case "groups"
        xpSMS.SetDetailsText "Manage contact groups as contacts can be grouped"
    Case "contacts"
        xpSMS.SetDetailsText "Manage contacts, cell numbers, emails etc"
    Case "msgs"
        xpSMS.SetDetailsText "Manage sms messages, inbox, outbox, drafts, scheduled etc"
    Case "phone"
        xpSMS.SetDetailsText "Browse your modem for contacts and messages"
    End Select
    Err.Clear
End Sub
Private Sub xpSMS_ItemClick(sGroup As String, sItemKey As String)
    On Error Resume Next
    ' when an item in navigation is clicked, do processes
    Dim lstSelected As ListItem
    Dim mResult As String
    Dim sKey As String
    Dim sOther As String
    tbListViewEdit.Visible = False
    sKey = MvField(sItemKey, 1, ",")
    sOther = Rest(sItemKey, 2, ",")
    cmdApply.Enabled = False
    Select Case LCase$(sKey)
    Case "potsmodem", "potsmodemtoserialport"
        ' view modem properties
        Dim vItems As Variant
        Dim vTemp As Variant
        Caption = "SMS Xpress Modem Properties: " & ProperCase$(sOther)
        ' Clear the ListView
        LstViewMakeHeadings lstReport, "Property,Value"
        lstReport.ListItems.Clear
        lstReport.View = lvwReport
        lstReport.Checkboxes = False
        ' Populate the ListView with the device's properties
        For Each vTemp In GetProperties(sKey, sOther)
            vItems = Split(vTemp, "^")
            lstReport.ListItems.Add(, , CStr(vItems(0))).SubItems(1) = vItems(1)
        Next
        LstViewAutoResize lstReport
        lstReport.Tag = "modems"
        ReadModem
        cmdApply.Enabled = True
        Err.Clear
        Exit Sub
    Case "msgtsk_forward"
        ' forward a message
        Select Case lstReport.Tag
        Case "draft", "outbox", "sentbox", "recycled", "inbox"
        Case Else
            Call MyPrompt("You have to be either accessing drafts, outbox, sentbox, inbox or recycled to be able to resend a message.", "o", "e", "Message Tasks")
            Err.Clear
            Exit Sub
        End Select
        ' get the selected item
        Set lstItem = lstReport.SelectedItem
        ' the user has not selected a group
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        ' get the name of the group
    Case "msgtsk_reply"
        ' reply to a message
        Select Case lstReport.Tag
        Case "inbox"
        Case Else
            Call MyPrompt("You have to be accessing the inbox to be able to reply to a message.", "o", "e", "Message Tasks")
            Err.Clear
            Exit Sub
        End Select
        ' get the selected item
        Set lstItem = lstReport.SelectedItem
        ' the user has not selected a group
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        ' reply to the sms
        With frmNewSMS
            .Caption = "Reply To SMS"
            .txtMsg.Text = ""
            .txtTo.Text = lstItem.SubItems(1)
            .txtReply.Text = lstItem.SubItems(2)
            .cmdSend.Enabled = True
            .cmdSaveDraft.Enabled = True
            .cmdOutbox.Enabled = True
            .Show vbModal
        End With
        lstReport.Tag = "newsms"
    Case "msgtsk_resend"
        ' resend a message
        Select Case lstReport.Tag
        Case "draft", "outbox", "sentbox", "recycled"
        Case Else
            Call MyPrompt("You have to be either accessing drafts, outbox, sentbox or recycled to be able to resend a message.", "o", "e", "Message Tasks")
            Err.Clear
            Exit Sub
        End Select
        ' get the selected item
        Set lstItem = lstReport.SelectedItem
        ' the user has not selected a group
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        ' send the message
        dbExchange_SendSMS lstItem.Text
        Select Case lstReport.Tag
        Case "draft"
            ViewMessages "draft"
        Case "outbox"
            ViewMessages "outbox"
        Case "sentbox"
            ViewMessages "sentbox"
        Case "recycled"
            ViewMessages "recycled"
        Case "inbox"
            ViewMessages "inbox"
        End Select
    Case "msgtsk_delete"
        ' delete a message
        Select Case lstReport.Tag
        Case "draft", "outbox", "sentbox", "recycled", "inbox"
        Case Else
            Call MyPrompt("You have to be either accessing drafts, outbox, sentbox, inbox or recycled to be able to delete a message.", "o", "e", "Message Tasks")
            Err.Clear
            Exit Sub
        End Select
        ' get the selected item
        Set lstItem = lstReport.SelectedItem
        ' the user has not selected a group
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        resp = MyPrompt("Are you sure that you want to recycle this message?", "yn", "q", "Delete Message: " & lstItem.Text)
        If resp = vbNo Then Exit Sub
        Call Me.dbExchange_Messages(Val(lstItem.Text), "", "", "", "Recycled", RecycleGroup)
        Select Case lstReport.Tag
        Case "draft"
            ViewMessages "draft"
        Case "outbox"
            ViewMessages "outbox"
        Case "sentbox"
            ViewMessages "sentbox"
        Case "recycled"
            ViewMessages "recycled"
        Case "inbox"
            ViewMessages "inbox"
        End Select
    Case "newmsg"
        ' send a new messag
        With frmNewSMS
            .Caption = "Send New SMS"
            .txtMsg.Text = ""
            .txtTo.Text = ""
            .txtReply.Text = ""
            .cmdSend.Enabled = False
            .cmdSaveDraft.Enabled = False
            .cmdOutbox.Enabled = False
            .Show vbModal
        End With
        lstReport.Tag = "newsms"
    Case "draft"
        ' view drafts
        ViewMessages "draft"
    Case "outbox"
        ' view outbox
        ViewMessages "outbox"
    Case "sentbox"
        ' view sentbox
        ViewMessages "sentbox"
    Case "recycled"
        ' view recycled
        ViewMessages "recycled"
    Case "inbox"
        ViewMessages "inbox"
    Case "groups_new"
        ' add a new group
AskGroup:
        sGroupName = InputBox("Please enter the new name of the group to add.", "New Group")
        If Len(sGroupName) = 0 Then Exit Sub
        If InStr(1, sGroupName, ",") > 0 Then
            resp = MyPrompt("The name of the group cannot have a comma.", "rc", "e", "Group Error")
            If resp = vbCancel Then Exit Sub
            GoTo AskGroup
        End If
        sGroupName = ProperCase(sGroupName)
        ' the group has been added
        If dbExchange_Groups(sGroupName, AddGroup) = True Then
            ' show all the groups
            xpSMS_ItemClick "", "groups_list"
        End If
        lstReport.Tag = "groups"
    Case "groups_delete"
        ' get the selected item
        Set lstItem = lstReport.SelectedItem
        ' the user has not selected a group
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        ' get the name of the group
        sGroupName = lstItem.Text
        ' ask the user if they want to delete the group and make no the default button
        resp = MsgBox("Are you sure that you want to delete this group?" & vbCr & vbCr & sGroupName, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
        ' if the user chooses no, stop the process
        If resp = vbNo Then Exit Sub
        ' delete the group from the database
        dbExchange_Groups sGroupName, DeleteGroup
        ' remove the group from the list
        lstReport.ListItems.Remove lstItem.Index
    Case "groups_sendmsg"
        ' send an sms to the group of contacts
        Set lstItem = lstReport.SelectedItem
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        SendSms_ToSomeOne "groupname", lstItem.Text
    Case "groups_list"
        Caption = "SMS Xpress: Groups"
        ' show all current available groups
        LstViewMakeHeadings lstReport, "Group Name"
        lstReport.Checkboxes = False
        lstReport.View = lvwReport
        lstReport.Tag = "groups"
        Dim rsGroups As dao.Recordset
        Set rsGroups = dbExchange.OpenRecordset("Groups")
        rsTot = rsGroups.RecordCount
        ProgBarInit progBar, rsTot
        For rsCnt = 1 To rsTot
            progBar.Value = rsCnt
            sGroupName = rsGroups!GroupName & ""
            lstReport.ListItems.Add , , sGroupName, "group", "group"
            DoEvents
            rsGroups.MoveNext
        Next
        progBar.Value = 0
        rsGroups.Close
        Set rsGroups = Nothing
        StatusMessage Me, lstReport.ListItems.Count & " group(s) listed."
    Case "contacts_new"
        ' add a new contact
        LstViewMakeHeadings lstReport, "Name,Cell Phone,Group"
        lstReport.ColumnHeaders(2).Alignment = lvwColumnRight
        lstReport.Checkboxes = True
        lstReport.Tag = "contacts"
        lstReport.Refresh
        Set lstItem = lstReport.ListItems.Add(, , "New Contact", "employee", "employee")
        lstItem.SubItems(1) = "0000000000"
        lstItem.SubItems(2) = "Unknown Group"
        lstReport.ListItems(1).Selected = True
        Call ListView_ScaleEdit(lstReport, tHt, tbListViewEdit)
        Call ListView_BeforeEdit(lstReport, tHt, tbListViewEdit)
    Case "contacts_delete"
        'delete selected contact
        Set lstItem = lstReport.SelectedItem
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        sGroupName = lstItem.SubItems(2)
        resp = MsgBox("Are you sure that you want to delete this contact?" & vbCr & vbCr & sGroupName, vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
        If resp = vbNo Then Exit Sub
        dbExchange_Contacts sGroupName, lstItem.SubItems(1), lstItem.SubItems(3), DeleteGroup
        lstReport.ListItems.Remove lstItem.Index
    Case "contacts_sendmsg"
        ' send message to selected contact
        Set lstItem = lstReport.SelectedItem
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        With frmNewSMS
            .Caption = "Send New SMS"
            .txtMsg.Text = ""
            .txtReply.Text = ""
            .txtTo.Text = lstItem.SubItems(1)
            .cmdSend.Enabled = False
            .cmdSaveDraft.Enabled = False
            .cmdOutbox.Enabled = False
            .Show vbModal
        End With
    Case "contacts_list"
        Dim rsContacts As dao.Recordset
        Dim rsIndex As Long
        Caption = "SMS Xpress: Contacts"
        ' show all current available contacts
        LstViewMakeHeadings lstReport, "Index,Cell Phone,Name,Group"
        lstReport.ColumnHeaders(2).Alignment = lvwColumnRight
        lstReport.Checkboxes = True
        lstReport.Tag = "contacts"
        Set rsContacts = dbExchange.OpenRecordset("contacts")
        rsIndex = 0
        Do Until rsContacts.EOF
            rsIndex = rsIndex + 1
            sGroup = rsContacts!GroupName & ""
            sName = rsContacts!Name & ""
            sCellPhone = rsContacts!cellphone & ""
            Set lstItem = lstReport.ListItems.Add(, , rsIndex, "employee", "employee")
            lstItem.SubItems(1) = sCellPhone
            lstItem.SubItems(2) = sName
            lstItem.SubItems(3) = sGroup
            DoEvents
            rsContacts.MoveNext
        Loop
        rsContacts.Close
        StatusMessage Me, lstReport.ListItems.Count & " contacts(s) listed."
    Case "phonecontacts_new"
        frmContacts.txtFullName.Text = ""
        frmContacts.txtNumber.Text = ""
        frmContacts.Show vbModal
    Case "phonecontacts_delete"
        Set lstSelected = lstReport.SelectedItem
        If TypeName(lstSelected) = "Nothing" Then
            Call MyPrompt("You have to select the contact to delete first.", "o", "e", "Delete Contact")
            Err.Clear
            Exit Sub
        End If
        resp = MsgBox("Are you sure that you want to delete this contact from the phonebook?" & vbCr & vbCr & lstSelected.SubItems(2), vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, "Confirm Delete")
        If resp = vbNo Then Exit Sub
        CheckSignal.Enabled = False
        Scheduler.Enabled = False
        Screen.MousePointer = vbHourglass
        StatusMessage Me, "Deleting contact from phonebook, please wait..."
        mResult = GSM.PhoneBook_DeleteEntry(Val(lstSelected.Text))
        Screen.MousePointer = vbDefault
        StatusMessage Me
        CheckSignal.Enabled = True
        Scheduler.Enabled = True
        If mResult = "OK" Then
            lstReport.ListItems.Remove lstSelected.Index
        Else
            Call MyPrompt("The selected contact could not be deleted.", "o", "e", "Delete Error")
        End If
    Case "phonecontacts_sendmsg"
        ' send message to selected contact
        Set lstItem = lstReport.SelectedItem
        If TypeName(lstItem) = "Nothing" Then Exit Sub
        With frmNewSMS
            .Caption = "Send New SMS"
            .txtMsg.Text = ""
            .txtReply.Text = ""
            .txtTo.Text = lstItem.SubItems(1)
            .cmdSend.Enabled = False
            .cmdSaveDraft.Enabled = False
            .cmdOutbox.Enabled = False
            .Show vbModal
        End With
    End Select
    Err.Clear
End Sub
Private Function dbExchange_Groups(sGroupName As String, GroupOperation As GroupsEnum) As Boolean
    On Error Resume Next
    If Len(sGroupName) = 0 Then Exit Function
    ' operate database record based on selected process
    Dim rsGroup As dao.Recordset
    ' open the groups table
    Set rsGroup = dbExchange.OpenRecordset("Groups")
    rsGroup.Index = "GroupName"
    rsGroup.Seek "=", sGroupName
    Select Case GroupOperation
    Case AddGroup
        If rsGroup.NoMatch = True Then
            rsGroup.AddNew
            rsGroup!GroupName = sGroupName
            rsGroup.Update
            dbExchange_Groups = True
        Else
            dbExchange_Groups = False
        End If
    Case DeleteGroup
        If rsGroup.NoMatch = False Then
            rsGroup.Delete
            dbExchange_Groups = True
        Else
            dbExchange_Groups = False
        End If
    End Select
    rsGroup.Close
    Set rsGroup = Nothing
    Err.Clear
End Function
Private Function dbExchange_Contacts(sName As String, sCellPhone As String, sGroup As String, GroupOperation As GroupsEnum) As Boolean
    On Error Resume Next
    Dim rsContacts As dao.Recordset
    ' open the contacts table, the name is the determinant
    Set rsContacts = dbExchange.OpenRecordset("Contacts")
    rsContacts.Index = "Cellphone"
    rsContacts.Seek "=", sCellPhone
    Select Case GroupOperation
    Case AddGroup
        If rsContacts.NoMatch = True Then
            rsContacts.AddNew
            rsContacts!Name = sName
            rsContacts!cellphone = sCellPhone
            rsContacts!GroupName = sGroup
            rsContacts.Update
        Else
            rsContacts.Edit
            rsContacts!Name = sName
            rsContacts!cellphone = sCellPhone
            rsContacts!GroupName = sGroup
            rsContacts.Update
        End If
        dbExchange_Contacts = True
    Case DeleteGroup
        If rsContacts.NoMatch = False Then
            rsContacts.Delete
            dbExchange_Contacts = True
        Else
            dbExchange_Contacts = False
        End If
    End Select
    rsContacts.Close
    Err.Clear
End Function
Private Function ComputerBook_Import(importFile As String) As Long
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
    Dim rsContacts As dao.Recordset
    Set rsContacts = dbExchange.OpenRecordset("contacts")
    rsContacts.Index = "Cellphone"
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
        rsContacts.Seek "=", cNumber
        Select Case rsContacts.NoMatch
        Case False
            rsContacts.Edit
            rsContacts!cellphone = cNumber
            rsContacts!Name = cName
            rsContacts.Update
        Case Else
            rsContacts.AddNew
            rsContacts!Name = cName
            rsContacts!cellphone = cNumber
            rsContacts.Update
            pWritten = pWritten + 1
        End Select
NextRow:
        DoEvents
    Next
    ComputerBook_Import = pWritten
    Err.Clear
End Function
Private Sub lstReport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Select Case lstReport.Tag
    Case "contacts", "phonebook_phone"
        Call ListView_AfterEdit(lstReport, tHt, tbListViewEdit)
        tHt = ListView_HitTest(lstReport, x, y)
        If tHt.lItem = -1 Then
        Else
            lstReport.ListItems(tHt.lItem + 1).Selected = True
        End If
    Case "draft", "outbox", "sentbox", "recycled", "inbox"
        If Button = 2 Then
            CboBoxLoadKeys dbExchange, "contacts", "name", cboGroups, , , , True
            ReloadMenu mnuContacts, cboGroups
            CboBoxLoadKeys dbExchange, "contacts", "groupname", cboGroups, , , , True
            ReloadMenu mnuGroups, cboGroups
            PopupMenu mnuFile
        End If
    End Select
    Err.Clear
End Sub
Private Sub tbListViewEdit_LostFocus()
    On Error Resume Next
    Dim sCols As String
    sCols = LstViewColNames(lstReport)
    SaveReg "column", "", App.Path, App.Title
    sCols = MvField(sCols, tHt.lSubItem + 1, ",")
    Select Case sCols
    Case "Name", "Cell Phone", "Group", "Cellphone No", "Full Name"
        SaveReg "column", sCols, "account", App.Title
    Case Else
        tbListViewEdit.Visible = False
    End Select
    Err.Clear
End Sub
Public Sub SendSMSOcx(ByVal sNumber As String, ByVal sMessage As String)
    On Error Resume Next
    Dim mAnswer As String
    mAnswer = GSM.SMS_Send(sNumber, sMessage)
    Select Case LCase$(mAnswer)
    Case "ok"
        Call frmSMSExchange.dbExchange_Messages(-1, sNumber, sMessage, Now(), "SentBox", AddGroup)
    Case Else
        Call frmSMSExchange.dbExchange_Messages(-1, sNumber, sMessage, Now(), "OutBox", AddGroup)
    End Select
    Err.Clear
End Sub
Public Function dbExchange_Messages(sID As Long, sTelephone As String, sContent As String, sProcessTime As String, sMsgType As String, GroupOperation As GroupsEnum) As Boolean
    On Error Resume Next
    Dim rsGroup As dao.Recordset
    ' open the messages table
    Set rsGroup = dbExchange.OpenRecordset("Messages")
    rsGroup.Index = "ID"
    rsGroup.Seek "=", sID
    Select Case GroupOperation
    Case AddGroup
        If rsGroup.NoMatch = True Then
            rsGroup.AddNew
            rsGroup!Telephone = sTelephone
            rsGroup!Content = sContent
            rsGroup!ProcessTime = IIf((sMsgType = "OutBox"), DateAdd("n", 1, sProcessTime), sProcessTime)
            rsGroup!msgType = sMsgType
            rsGroup.Update
            dbExchange_Messages = True
        Else
            dbExchange_Messages = False
        End If
    Case DeleteGroup
        If rsGroup.NoMatch = False Then
            rsGroup.Delete
            dbExchange_Messages = True
        Else
            dbExchange_Messages = False
        End If
    Case RecycleGroup
        If rsGroup.NoMatch = False Then
            rsGroup.Edit
            rsGroup!msgType = sMsgType
            rsGroup.Update
            dbExchange_Messages = True
        Else
            dbExchange_Messages = True
        End If
    End Select
    rsGroup.Close
    Set rsGroup = Nothing
    Err.Clear
End Function
Private Sub RefreshModems()
    On Error Resume Next
    ' get the names of all available modems in the computer
    xpSMS.DisableUpdates True
    xpSMS.ClearGroup "modems"
    Dim DevicesNames As Variant
    Dim Device As Variant
    Dim TempDevice As Variant
    Dim NumOfDevices As Integer
    Dim Devices As Variant
    Dim tmpDevice As String
    ' This array contains the Computer System Hardware Classes names, we will only look at the modems
    DevicesNames = Array("Win32_POTSModem", "Win32_POTSModemToSerialPort")
    ' Find the number of hardware classes
    NumOfDevices = UBound(DevicesNames)
    ' Find all the hardware devices
    For Each Device In DevicesNames
        ' Make sure that the operating system can process other events
        DoEvents
        tmpDevice = Right$(Device, Len(Device) - 6)
        Devices = GetDevice(CStr(Device))
        For Each TempDevice In Devices
            ' Add the device name to the group
            'xpSMS.AddItem "modems", tmpDevice & "," & CStr(TempDevice), CStr(TempDevice), "modem"
            treePhone.Nodes.Add "modems", tvwChild, tmpDevice & "," & CStr(TempDevice), CStr(TempDevice), "modem", "modem"
        Next
    Next
    xpSMS.DisableUpdates False
    Err.Clear
End Sub
Private Function GetDevice(DeviceName As String) As Variant
    On Error Resume Next
    ' In this function we will get the devices referring to the given class name
    Dim DeviceSet As SWbemObjectSet
    Dim Device As SWbemObject
    Dim sTemp As String
    ' Set the SWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf(DeviceName)
    ' Get the devices captions
    For Each Device In DeviceSet
        sTemp = sTemp & Device.Caption & "|"
    Next
    ' Remove the '|' character at the end of the string
    If Right$(sTemp, 1) = "|" Then sTemp = Left$(sTemp, Len(sTemp) - 1)
    ' Return an array (variant) with the devices captions
    GetDevice = Split(sTemp, "|")
    Err.Clear
End Function
Private Function GetProperties(sInstance As String, sDevice As String) As Variant
    On Error Resume Next
    ' This function returns all the properties of a specific device
    Dim DeviceSet As SWbemObjectSet
    Dim Device As SWbemObject
    Dim vTemp As Variant
    Dim sTemp As String
    ' Set theSWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf("Win32_" & sInstance)
    For Each Device In DeviceSet
        ' Check if the current device in the chosen device
        If LCase$(Device.Caption) = LCase$(sDevice) Then
            ' Get all the properties of the chosen device
            For Each vTemp In Device.Properties_
                If vTemp <> "" And vTemp <> vbNull Then
                    ' Add the property name and its value to the temporary string
                    sTemp = sTemp & vTemp.Name & "^" & vTemp & "|"
                End If
            Next
            ' Remove the '|' character at the end of the string
            If Right$(sTemp, 1) = "|" Then
                sTemp = Left$(sTemp, Len(sTemp) - 1)
            End If
        End If
    Next
    ' Return an array containing the device properties
    GetProperties = Split(sTemp, "|")
    Err.Clear
End Function
Private Sub ReadModem()
    On Error Resume Next
    ' read modem properties from listview
    Dim lPos As Long
    Dim myRow() As String
    Dim mSignal As String
    ' find the name of the modem and display it
    lPos = LstViewFindItem(lstReport, "Caption")
    If lPos > 0 Then
        myRow = LstViewGetRow(lstReport, lPos)
        lblModem.Caption = myRow(2)
    Else
        lblModem.Caption = ""
    End If
    ' find the port that the modem is connected to
    lPos = LstViewFindItem(lstReport, "AttachedTo")
    If lPos > 0 Then
        myRow = LstViewGetRow(lstReport, lPos)
        txtPort.Text = ExtractNumbers(myRow(2))
    Else
        txtPort.Text = ""
    End If
    ' find the maximum connection rate as set as settings
    '"460800,n,8,1"
    lPos = LstViewFindItem(lstReport, "maxbaudratetoserialport")
    If lPos > 0 Then
        myRow = LstViewGetRow(lstReport, lPos)
        txtMaxSpeed.Text = myRow(2)
        txtSettings.Text = myRow(2) & ",n,8,1"
    Else
        txtSettings.Text = ""
        txtMaxSpeed.Text = ""
    End If
    CheckSignal.Enabled = False
    Scheduler.Enabled = False
    StatusMessage Me, "Connecting to modem..."
    If GSM.Connect(txtPort.Text, txtMaxSpeed.Text) = "OK" Then
        StatusMessage Me, "Reading the message centre number..."
        txtMCN.Text = GSM.SMS_CentreNumber(ReadCentreNumber)
        StatusMessage Me, "Checking the signal quality..."
        mSignal = GSM.SignalQualityMeasure
        progSignal.Value = Val(mSignal)
        StatusMessage Me, lblModem.Caption & " (" & mSignal & "% Signal)", 2
        frmSYS_Splash.lblLicenseTo.Caption = "Setting new message indicators on..."
        GSM.SMS_NewMessageIndicate False
        StatusMessage Me, "Reading imei number of modem..."
        txtIMEI.Text = GSM.ModemSerialNumber
        StatusMessage Me, "Setting sms mode to text format..."
        Call GSM.SMS_MessageFormat(TextFormat)
        StatusMessage Me, "Setting sms preferred storage area..."
        Call GSM.SMS_MemoryStorage(MobileEquipmentMemory)
        StatusMessage Me, "Setting / returning phonebook preferred storage area..."
        Call GSM.PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
        Call GSM.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
        StatusMessage Me, GSM.PhoneBook_Memory & " Phonebook, Capacity " & GSM.PhoneBook_Capacity & ", Used " & GSM.PhoneBook_Used, 3
        StatusMessage Me
    End If
    CheckSignal.Enabled = True
    Scheduler.Enabled = True
    StatusMessage Me
    Err.Clear
End Sub
Public Sub dbExchange_SendSMS(ByVal msgNumber As String, Optional OnReport As Boolean = True)
    On Error Resume Next
    ' send a particular msg from saved records in the database
    Dim rsMessages As dao.Recordset
    Dim sTelephone As String
    Dim sContent As String
    Dim sMsgType As String
    Dim msgPos As Long
    Dim mAnswer As String
    Set rsMessages = dbExchange.OpenRecordset("select * from Messages where ID = " & msgNumber & ";")
    rsMessages.MoveLast
    rsMessages.MoveFirst
    If Err.Number = 3021 Then GoTo ExitSender
    sTelephone = rsMessages!Telephone & ""
    sContent = rsMessages!Content & ""
    sMsgType = rsMessages!msgType & ""
    mAnswer = GSM.SMS_Send(sTelephone, sContent)
    Select Case mAnswer
    Case "OK"
        ' the message was sent successfully
        ' check the source of the message
        Select Case sMsgType
        Case "SentBox", "Inbox"
            ' the message has been resent, thus being forwarded, then add a new message
            Call dbExchange_Messages(-1, sTelephone, sContent, Now(), "SentBox", AddGroup)
        Case "Draft", "OutBox"
            ' the message is from draft/outbox, update as sent
            rsMessages.Edit
            rsMessages!msgType = "SentBox"
            rsMessages!ProcessTime = Now()
            rsMessages.Update
            If OnReport = True Then
                msgPos = LstViewFindItem(lstReport, msgNumber, search_Text, search_Whole)
                If msgPos > 0 Then
                    lstReport.ListItems.Remove msgPos
                End If
            End If
        End Select
    Case Else
        ' the message could not be sent
        ' check the source of the message
        Select Case sMsgType
        Case "SentBox", "Inbox"
            ' the message has been resent, thus being forwarded, then add a new message
            Call dbExchange_Messages(-1, sTelephone, sContent, Now(), "OutBox", AddGroup)
        Case "Draft"
            ' the message is from draft/outbox, save to outbox
            rsMessages.Edit
            rsMessages!msgType = "OutBox"
            rsMessages!ProcessTime = DateAdd("n", 1, Now())
            rsMessages.Update
            If OnReport = True Then
                msgPos = LstViewFindItem(lstReport, msgNumber, search_Text, search_Whole)
                If msgPos > 0 Then
                    lstReport.ListItems.Remove msgPos
                End If
            End If
        Case "OutBox"
            ' the message is from draft/outbox, save to outbox
            rsMessages.Edit
            rsMessages!msgType = "OutBox"
            rsMessages!ProcessTime = DateAdd("n", 1, Now())
            rsMessages.Update
        End Select
    End Select
ExitSender:
    rsMessages.Close
    Err.Clear
End Sub
Public Sub ViewMessages(ByVal strFolder As String)
    On Error Resume Next
    ' view messages based on selected msgtype
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsMessages As dao.Recordset
    Dim sID As String
    Dim sTelephone As String
    Dim sProcessTime As String
    Dim sContent As String
    LstViewMakeHeadings lstReport, "Msg ID,Cellphone No.,Contents,Time"
    lstReport.ListItems.Clear
    qrySql = "select id,telephone,content,processtime from Messages where msgtype = '" & strFolder & "' order by processtime desc;"
    Caption = "SMS Xpress: " & ProperCase(strFolder) & " Messages"
    ' show all current available draft messages
    lstReport.Checkboxes = False
    Set rsMessages = dbExchange.OpenRecordset(qrySql)
    rsMessages.MoveLast
    rsTot = rsMessages.RecordCount
    ProgBarInit progBar, rsTot
    rsMessages.MoveFirst
    For rsCnt = 1 To rsTot
        progBar.Value = rsCnt
        sID = rsMessages!ID & ""
        sTelephone = rsMessages!Telephone & ""
        sContent = rsMessages!Content & ""
        sProcessTime = rsMessages!ProcessTime & ""
        Set lstItem = lstReport.ListItems.Add(, , sID, strFolder, strFolder)
        lstItem.SubItems(1) = sTelephone
        lstItem.SubItems(2) = sContent
        lstItem.SubItems(3) = sProcessTime
        DoEvents
        rsMessages.MoveNext
    Next
    rsMessages.Close
    progBar.Value = 0
    LstViewAutoResize lstReport
    StatusMessage Me, lstReport.ListItems.Count & " message(s) listed."
    lstReport.Tag = LCase$(strFolder)
    Err.Clear
End Sub
Public Sub SendSms_ToSomeOne(ByVal sFieldName As String, ByVal sGroupName As String, Optional msgNumber As Long = -1)
    On Error Resume Next
    ' create a new message for sending to selected group/contact
    Dim sName As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsContacts As dao.Recordset
    Set rsContacts = dbExchange.OpenRecordset("select Cellphone from contacts where " & sFieldName & " = '" & sGroupName & "';")
    sGroupName = ""
    rsContacts.MoveLast
    rsTot = rsContacts.RecordCount
    rsContacts.MoveFirst
    For rsCnt = 1 To rsTot
        sName = rsContacts!cellphone & ""
        sGroupName = sGroupName & sName & ","
        rsContacts.MoveNext
    Next
    rsContacts.Close
    sGroupName = RemDelim(sGroupName, ",")
    With frmNewSMS
        .Caption = "Send New SMS"
        .txtMsg.Text = ""
        .txtReply.Text = ""
        .txtTo.Text = sGroupName
        .cmdSend.Enabled = False
        .cmdSaveDraft.Enabled = False
        .cmdOutbox.Enabled = False
        If msgNumber > 0 Then
            Set rsContacts = dbExchange.OpenRecordset("select Content from Messages where id = " & msgNumber)
            rsContacts.MoveLast
            .txtMsg.Text = rsContacts!Content & ""
        End If
        .Show vbModal
    End With
    Err.Clear
End Sub
Private Function dbExchange_Gadgets(sImei As String, sCapacity As String, GroupOperation As GroupsEnum) As Boolean
    On Error Resume Next
    If Len(sName) = 0 And Len(sCellPhone) = 0 Then Exit Function
    Dim rsGadgets As dao.Recordset
    ' open the gadgets table, the imei is the determinant
    Set rsGadgets = dbExchange.OpenRecordset("Gadgets")
    rsGadgets.Index = "IMEI"
    rsGadgets.Seek "=", sImei
    Select Case GroupOperation
    Case AddGroup
        If rsGadgets.NoMatch = True Then
            rsGadgets.AddNew
            rsGadgets!imei = sImei
            rsGadgets!capacity = Val(sCapacity)
            rsGadgets.Update
        Else
            rsGadgets.Edit
            rsGadgets!imei = sImei
            rsGadgets!capacity = Val(sCapacity)
            rsGadgets.Update
        End If
        dbExchange_Gadgets = True
    Case DeleteGroup
        If rsGadgets.NoMatch = False Then
            rsGadgets.Delete
            dbExchange_Gadgets = True
        Else
            dbExchange_Gadgets = False
        End If
    End Select
    rsGadgets.Close
    Err.Clear
End Function
Private Sub txtIMEI_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtIMEI, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtMaxSpeed_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtMaxSpeed, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtSettings_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtSettings, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtPort_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtPort, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtMCN_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, txtMCN, KeyCode, Shift
    Err.Clear
End Sub
Private Sub tbListViewEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Intel.Sense App.Path, tbListViewEdit, KeyCode, Shift
    Err.Clear
End Sub
Private Sub txtIMEI_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtIMEI
    Err.Clear
End Sub
Private Sub txtMaxSpeed_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtMaxSpeed
    Err.Clear
End Sub
Private Sub txtSettings_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtSettings
    Err.Clear
End Sub
Private Sub txtPort_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtPort
    Err.Clear
End Sub
Private Sub txtMCN_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtMCN
    Err.Clear
End Sub
Private Sub tbListViewEdit_GotFocus()
    On Error Resume Next
    TextBoxHiLite tbListViewEdit
    Err.Clear
End Sub
Private Sub txtIMEI_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtIMEI.Text) = 0 Then Exit Sub
    txtIMEI.Text = ProperCase(txtIMEI.Text)
    Err.Clear
End Sub
Private Sub txtMaxSpeed_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtMaxSpeed.Text) = 0 Then Exit Sub
    txtMaxSpeed.Text = ProperCase(txtMaxSpeed.Text)
    Err.Clear
End Sub
Private Sub txtSettings_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtSettings.Text) = 0 Then Exit Sub
    txtSettings.Text = ProperCase(txtSettings.Text)
    Err.Clear
End Sub
Private Sub txtPort_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtPort.Text) = 0 Then Exit Sub
    txtPort.Text = ProperCase(txtPort.Text)
    Err.Clear
End Sub
Private Sub txtMCN_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(txtMCN.Text) = 0 Then Exit Sub
    txtMCN.Text = ProperCase(txtMCN.Text)
    Err.Clear
End Sub
Private Sub tbListViewEdit_Validate(Cancel As Boolean)
    On Error Resume Next
    If Len(tbListViewEdit.Text) = 0 Then Exit Sub
    tbListViewEdit.Text = ProperCase(tbListViewEdit.Text)
    Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FormMoveToNextControl KeyAscii
    Err.Clear
End Sub
Private Sub ExportContactsToFile()
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim bSelected As Boolean
    Dim exportFile As String
    Dim rsLine As String
    exportFile = App.Path & "\copy contacts.csv"
    If FileExists(exportFile) = True Then Kill exportFile
    FileUpdate exportFile, "Index,Cellphone No,Full Name", "a"
    rsTot = lstReport.ListItems.Count
    ProgBarInit progBar, rsTot
    For rsCnt = 1 To rsTot
        DoEvents
        progBar.Value = rsCnt
        bSelected = lstReport.ListItems(rsCnt).Checked
        If bSelected = False Then GoTo NextRow
        spLine = LstViewGetRow(lstReport, rsCnt)
        rsLine = spLine(1) & "," & spLine(2) & "," & spLine(3)
        FileUpdate exportFile, rsLine, "a"
NextRow:
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Function ComputerBook_Export() As Boolean
    On Error Resume Next
    Dim rsContacts As dao.Recordset
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim exportFile As String
    Dim rsLine As String
    Dim cName As String
    Dim cNumber As String
    ' open the contacts table, the name is the determinant
    Set rsContacts = dbExchange.OpenRecordset("Contacts")
    exportFile = App.Path & "\copy contacts.csv"
    If FileExists(exportFile) = True Then Kill exportFile
    FileUpdate exportFile, "Index,Cellphone No,Full Name", "a"
    rsContacts.MoveLast
    rsTot = rsContacts.RecordCount
    rsContacts.MoveFirst
    ProgBarInit progBar, rsTot
    For rsCnt = 1 To rsTot
        progBar.Value = 0
        cName = rsContacts!Name & ""
        cNumber = rsContacts!cellphone & ""
        rsLine = rsCnt & "," & cNumber & "," & cName
        FileUpdate exportFile, rsLine, "a"
        DoEvents
        rsContacts.MoveNext
    Next
    rsContacts.Close
    ComputerBook_Export = True
    Err.Clear
End Function
