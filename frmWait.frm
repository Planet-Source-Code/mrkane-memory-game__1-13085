VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmWait.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please stand by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
