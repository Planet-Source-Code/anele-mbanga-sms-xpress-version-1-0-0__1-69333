VERSION 5.00
Begin VB.Form frmSYS_Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   8505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   8265
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":1C7A
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7200
         TabIndex        =   3
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Warning: This is opensource, knowledge belongs to the world."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   4470
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7380
         TabIndex        =   4
         Top             =   2880
         Width           =   525
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   675
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmSYS_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error Resume Next
    lblLicenseTo.Caption = ""
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.ProductName
    lblCopyright.Caption = "Copyright: " & App.LegalCopyright
    Err.Clear
End Sub
