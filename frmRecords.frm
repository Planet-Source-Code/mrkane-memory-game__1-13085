VERSION 5.00
Begin VB.Form frmRecords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Memory Game Records"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmRecords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "&Remove All"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstRecords 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   5520
      MouseIcon       =   "frmRecords.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "frmRecords.frx":015E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdRemoveAll_Click()
    On Error Resume Next
    If MsgBox("You are about to remove all existing records." & Chr(10) & Chr(10) & "Click 'OK' to continue." & Chr(10) & "Click 'Cancel' to abort.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then Exit Sub
    DeleteSetting App.Title, "Records"
    lstRecords.Clear
    lstRecords.AddItem "Records deleted."
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Records As Variant, Record As Integer, Name As String, Record2 As Integer, Name2 As String
    Records = GetAllSettings(App.Title, "Records")
    
    lstRecords.Clear
    lstRecords.AddItem "©V.E.K. Software, 2000"
    For i = 1 To 576
            Record = GetSetting(App.Title, "Records", i, 9999)
            Name = GetSetting(App.Title, "Records", i & "_NAME", "")
            Record2 = GetSetting(App.Title, "Records", i & "_FLIP", 9999)
            Name2 = GetSetting(App.Title, "Records", i & "_FLIP_NAME", "")
            If Record <> 9999 Or Record2 <> 9999 Or Name <> "" Or Name2 <> "" Then
            lstRecords.AddItem " "
            lstRecords.AddItem i & " pieces"
            lstRecords.AddItem "---------------"
            End If
            If Record <> 9999 And Name <> "" Then
                lstRecords.AddItem "Time: " & Name & ", with " & Record & " seconds."
            End If
            If Recor2 <> 9999 And Name2 <> "" Then
                lstRecords.AddItem "Flips: " & Name2 & ", with " & Record2 & " flips."
            End If
    Next i
End Sub
