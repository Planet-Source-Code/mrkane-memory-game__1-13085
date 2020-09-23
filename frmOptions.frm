VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlStuff 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   2355
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
      Begin VB.ComboBox cboCollection2 
         Height          =   360
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ImageCombo cboCollection 
         Height          =   330
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         ImageList       =   "imlIcons"
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select a theme of icons to use:"
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   120
         Width           =   3255
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmOptions.frx":000C
         Picture         =   "frmOptions.frx":015E
         Top             =   240
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "&Abort"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   1931
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
      Begin VB.ComboBox cboRows 
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboCols 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select the number of cards:"
         Height          =   240
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   120
         Width           =   2985
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmOptions.frx":05A0
         Picture         =   "frmOptions.frx":06F2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   4080
      MouseIcon       =   "frmOptions.frx":0B34
      MousePointer    =   99  'Custom
      Picture         =   "frmOptions.frx":0C86
      Top             =   2760
      Width           =   480
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SaveSettings()
    SaveSetting App.Title, "Settings", "Cols", cboCols
    SaveSetting App.Title, "Settings", "Rows", cboRows
    SaveSetting App.Title, "Settings", "Collection", cboCollection.SelectedItem.Text
    
End Sub
Sub LoadSettings()
    Dim Coll As String
    cboCols = GetSetting(App.Title, "Settings", "Cols", 4)
    cboRows = GetSetting(App.Title, "Settings", "Rows", 4)
    Coll = GetSetting(App.Title, "Settings", "Collection", "flags")
    If Dir(App.Path & "\icons\" & Coll, vbDirectory) <> "" Then
        For i = 1 To cboCollection.ComboItems.Count
            If cboCollection.ComboItems.Item(i).Text = Coll Then
                cboCollection.ComboItems.Item(i).Selected = True
            End If
        Next i
        'cboCollection.GetFirstVisible
    Else
        If Dir(App.Path & "\icons\flags", vbDirectory) = "" Then
            MsgBox "Default icon theme could not be found.", 16, "Error"
            cboCollection.ComboItems.Item(1).Selected = True
            Exit Sub
        Else
            For i = 1 To cboCollection.ComboItems.Count
                If cboCollection.ComboItems.Item(i).Text = "flags" Then
                    cboCollection.ComboItems.Item(i).Selected = True
                End If
            Next i
        End If
        'MsgBox "Missing icon collection. Selecting default(flags)...", vbExclamation, "Warning"
    End If
    
End Sub
Sub FillCR()
    cboCols.Clear
    cboRows.Clear
    For i = 1 To 25
        cboCols.AddItem i
        cboRows.AddItem i
    Next i
End Sub
Sub FillCollection()
    'On Error Resume Next
    Dim SubFolder As Folder, Cntr As Integer
    
    For Each SubFolder In FSO.GetFolder(App.Path & "\icons").SubFolders
        Cntr = Cntr + 1
        a = SubFolder.Path & "\" & Dir(SubFolder.Path & "\*.ICB")
        Me.imlIcons.ListImages.Add , , LoadPicture(a)
        cboCollection.ComboItems.Add , UCase(SubFolder.Name), LCase(SubFolder.Name), Cntr
    Next
End Sub

Private Sub cmdAbort_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If (cboCols * cboRows) Mod 2 <> 0 Then
        MsgBox "The number of rows and collumns you choose must result in an even number ( to make pairs )!", 16, "Error"
        Exit Sub
    End If
    SaveSettings
    'UpdateIT
    Unload Me
End Sub

Private Sub Form_Load()
    FillCollection
    FillCR
    LoadSettings
End Sub
