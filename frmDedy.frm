VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDedy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Memory Game"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDedy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlDedy 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   3625
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Label lblDedy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please stand, by, loading dedication file data..."
         Height          =   1785
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4545
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Play"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   360
      MouseIcon       =   "frmDedy.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmDedy.frx":0594
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Memory Game!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3960
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmDedy.frx":09D6
      MousePointer    =   99  'Custom
      Picture         =   "frmDedy.frx":0B28
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then End
End Sub


