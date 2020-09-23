VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlSplash 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   4260
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
      BevelInner      =   2
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   2940
         MouseIcon       =   "frmSplash.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmSplash.frx":0152
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   180
         MouseIcon       =   "frmSplash.frx":0594
         MousePointer    =   99  'Custom
         Picture         =   "frmSplash.frx":06E6
         Top             =   120
         Width           =   480
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
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
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
         TabIndex        =   4
         Top             =   960
         Width           =   1980
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
         TabIndex        =   3
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblInfo 
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
         Left            =   780
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VB.Line lnDeco 
         X1              =   180
         X2              =   3420
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE WAIT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   1800
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblInfo(0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    HCenterControl lblInfo(0)
    HCenterControl lblInfo(2)
    HCenterControl lblInfo(3)
    HCenterControl lblStatus
    lblStatus = "LOADING GAME" & Chr(10) & FSO.GetFolder(App.Path & "\icons").SubFolders.Count & " icon theme(s)"
End Sub

