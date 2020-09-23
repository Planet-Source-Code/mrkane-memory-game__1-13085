VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlAbout 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   3201
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memory Game"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   765
         TabIndex        =   5
         Top             =   120
         Width           =   2085
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version x.x.x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Written by Mr. Kane"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   960
         Width           =   1980
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "©V.E.K. Software, 2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   180
         MouseIcon       =   "frmAbout.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":015E
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   2940
         MouseIcon       =   "frmAbout.frx":05A0
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":06F2
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblInfo(0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    HCenterControl cmdOK
    HCenterControl lblInfo(0)
    HCenterControl lblInfo(2)
    HCenterControl lblInfo(3)
End Sub

