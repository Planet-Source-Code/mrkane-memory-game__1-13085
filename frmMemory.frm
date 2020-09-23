VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmMemory 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Game"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel picPanel 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "©V.E.K. Software"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   2
      Alignment       =   6
      Begin VB.Image imgPiece 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMemory.frx":0442
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
   End
   Begin MCI.MMControl mciWav 
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   -120
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1296
      _Version        =   327680
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WAVEAUDIO"
      FileName        =   ""
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5280
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Timer tmrTime 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.ListBox lstBacks 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox lstIcons 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.ListBox lstIcons 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Line lnDeco 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblBestFlips 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Best Flips:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4575
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblBestTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Best Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblFlips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flips:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   390
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuRecords 
         Caption         =   "&View records..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpMe 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub PSND()
    'Exit Sub
    mciWav.Command = "PREV"
    mciWav.Command = "PLAY"
End Sub
Sub CheckDedication()
    Dim TMP1 As String
    If Dir(App.Path & "\memory.vsd") = "" Then Exit Sub
    
    Open App.Path & "\memory.vsd" For Random As #1 Len = 5000
    Get #1, 6, TMP1
    If TMP1 <> "4387348745853274538434583754543545" Then Exit Sub
    Get #1, 1, TMP1
    frmDedy.lblDedy = TMP1
    frmDedy.Show vbModal

    
    
    Close #1
End Sub

Private Sub DrawBackGround()
    'Exit Sub
    lblTime.Visible = False
    lblFlips.Visible = False
    lblBestTime.Visible = False
    
    lblBestFlips.Visible = False
    Me.lnDeco.Visible = False
    Const intBLUESTART% = 0
    Const intBLUEEND% = 255
    Const intBANDHEIGHT% = 1
    Const intSHADOWSTART% = 8
    Const intSHADOWCOLOR% = 0
    Const intTEXTSTART% = 4
    Const intTEXTCOLOR% = 15

    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer

    '
    'Get system values for height and width
    '
    intFormHeight = ScaleHeight
    intFormWidth = ScaleWidth

    '
    'Calculate step size and blue start value
    '
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART

    '
    'Paint blue screen
    '
    For intY = 0 To intFormHeight / 2 Step intBANDHEIGHT
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur * 2), BF
        sngBlueCur = sngBlueCur + sngBlueStep
        
        'If Right(Str(intY), 1) = "0" Then DoEvents
    Next intY
    For intY = intFormHeight / 2 To intFormHeight Step intBANDHEIGHT
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur * 2), BF
        sngBlueCur = sngBlueCur - sngBlueStep
        'If Right(Str(intY), 1) = "0" Then DoEvents
    Next intY
    
    lblTime.Visible = True
    lblFlips.Visible = True
    lblBestTime.Visible = True
    lblBestFlips.Visible = True
    Me.lnDeco.Visible = True
    
    
    
    
    
    
    '
    'Print 'shadowed' appname
    '
    
    Exit Sub
    CurrentX = intSHADOWSTART
    CurrentY = intSHADOWSTART
    ForeColor = QBColor(intSHADOWCOLOR)
    Print Caption
    CurrentX = intTEXTSTART
    CurrentY = intTEXTSTART
    ForeColor = QBColor(intTEXTCOLOR)
    Print Caption
End Sub
Sub CheckForWin()
    On Error GoTo ErrHandling
    Dim Win As Boolean, NewMaster As String, Response As Integer
    Win = True

    For i = 0 To imgPiece.Count - 1
        If imgPiece(i).Tag <> "DONE" Then

            Win = False
            Exit Sub
        End If
    Next i
    tmrTime.Enabled = False
    MsgBox "Very well." & Chr(10) & "You have completed the game." & Chr(10) & "And it took you ... " & Game.Time & " second(s) to do it!" & Chr(10) & "It took you " & Game.Moves & " moves.", vbExclamation, "Game Over"
    If Game.Time < Game.Record.Time Then
        MsgBox "You have set a new record! Congratulations!" & Chr(10) & "A new best time, eh!", vbExclamation, "TimeMachine"
EnterName:
        NewMaster = InputBox("Please enter your name (for the record):", "New Record")
        If NewMaster = "" Then
            MsgBox "You must enter your name.", vbCritical, "Warning"
            GoTo EnterName
        End If
            
        SaveSetting App.Title, "Records", Game.Total & "_NAME", NewMaster
        SaveSetting App.Title, "Records", Game.Total, Game.Time
    End If
    If Game.Moves < Game.Record.Flips Then
        MsgBox "You have set a new record! Congratulations!" & Chr(10) & "Less moves than the last one, eh!", vbExclamation, "FlipMaster"
        If NewMaster = "" Then GoTo EnterName2
        GoTo CheckIt
EnterName2:
        NewMaster = InputBox("Please enter your name (for the record):", "New Record")
CheckIt:
        If NewMaster = "" Then
            MsgBox "You must enter your name.", vbCritical, "Warning"
            GoTo EnterName2
        End If
            
        SaveSetting App.Title, "Records", Game.Total & "_FLIP_NAME", NewMaster
        SaveSetting App.Title, "Records", Game.Total & "_FLIP", Game.Moves
    End If
    
    StartGame
    Exit Sub
ErrHandling:
    
    Response = MsgBox("An error has occured in the CheckForWin() routine. Below is the number and description of the error:" & Chr(10) & Err.Number & ":" & Err.Description, vbCritical + vbAbortRetryIgnore, "Error")
    Select Case Response
        Case vbRetry
            Resume
        Case vbAbort
            Exit Sub
        Case vbIgnore
            Resume Next
    End Select

End Sub
Sub LoadLists()
    On Error GoTo ErrHandling
    Dim Counter As Integer, FileName As String, Response As Integer, ICNMODE As Integer
    lstBacks.Clear
    lstIcons(0).Clear
    lstIcons(1).Clear
    FileName = Dir(App.Path & "\icons\" & Game.Collection & "\*.ICB")
    While FileName <> ""
        lstBacks.AddItem App.Path & "\icons\" & Game.Collection & "\" & FileName
        FileName = Dir
        
    Wend
    
    lstIcons(0).Clear
    FileName = Dir(App.Path & "\icons\" & Game.Collection & "\*.ICO")
    While FileName <> ""
        
        If FileName <> "" Then lstIcons(0).AddItem App.Path & "\icons\" & Game.Collection & "\" & FileName
        FileName = Dir
        
    Wend

Check:
    If lstIcons(0).ListCount < (Game.Total / 2) Then '
ShuffleIt:
        lstIcons(0).ListIndex = Int(Rnd * (lstIcons(0).ListCount - 1))
        'lstIcons(0).AddItem lstIcons(0)
        lstIcons(0).AddItem lstIcons(0)

        GoTo Check
    End If

    For i = 0 To (Game.Total - 1) / 2
        
        
        ICNMODE = 1
        If Dir(App.Path & "\icons\" & Game.Collection & "\mgame.ITP") <> "" Then
            
            Open App.Path & "\icons\" & Game.Collection & "\mgame.ITP" For Input As #1
            Input #1, ICNmodetmp
            ICNMODE = Val(ICNmodetmp)
            Close #1
        End If
        
        If ICNMODE = 1 Then
            lstIcons(0).ListIndex = Rnd * (lstIcons(0).ListCount - 1)
            lstIcons(1).AddItem lstIcons(0).List(lstIcons(0).ListIndex)
            lstIcons(1).AddItem lstIcons(0).List(lstIcons(0).ListIndex)
        Else
            lstIcons(0).ListIndex = Rnd * (lstIcons(0).ListCount - 1)
            'flstIcons(0).ListIndex = i
            lstIcons(1).AddItem lstIcons(0).List(i)
            lstIcons(1).AddItem lstIcons(0).List(i)
        End If
        
        lstIcons(0).RemoveItem lstIcons(0).ListIndex
    Next i
    Exit Sub
ErrHandling:
    
    Response = MsgBox("An error has occured in the LoadLists() routine. Below is the number and description of the error:" & Chr(10) & Err.Number & ":" & Err.Description, vbCritical + vbAbortRetryIgnore, "Error")
    Select Case Response
        Case vbRetry
            Resume
        Case vbAbort
            Exit Sub
        Case vbIgnore
            Resume Next
    End Select


    

End Sub
Sub StartGame()
    On Error GoTo ErrHandling

    Dim NewWidth As Integer, NewHeight As Integer, Response As Integer
    'Me.Enabled = False
    UpdateIt
    LoadLists
    
    
    Game.Time = 0
    Game.Moves = 0
    lblBestTime = "Best time:" & Game.Record.Time & " second(s), " & Game.Record.TimeMaster
    lblBestFlips = "Best flips:" & Game.Record.Flips & " moves, " & Game.Record.Flipper
    lblTime = "Begin when ready."
    lblFlips = Game.Rows & " x " & Game.Cols & " pieces board"

    tmrTime.Enabled = False
    picPanel.Visible = False
    DoEvents
    Game.Opened = False
    Game.LastOpen1 = -1
    Game.LastOpen2 = -1
    Game.PicBack = lstBacks.List(Rnd * (lstBacks.ListCount - 1))
    'Removes any "leftovers"
    For i = 1 To imgPiece.Count - 1
        Unload imgPiece(i)
    Next i
    
    'Loads pieces and places them
    
    For i = 0 To Game.Total - 1
        If i > 0 Then Load imgPiece(i)
        imgPiece(i).BorderStyle = 0
        
        imgPiece(i).Enabled = True
        'Loads the actual content of the card
        
        lstIcons(1).ListIndex = Rnd * (lstIcons(1).ListCount - 1)
         imgPiece(i).Tag = lstIcons(1)
        lstIcons(1).RemoveItem lstIcons(1).ListIndex
        
        'Shows the back of the card
        imgPiece(i).Picture = LoadPicture(Game.PicBack)
        imgPiece(i).Visible = True
        
        'Places the cards
        If i > 0 Then
            If i Mod Game.Cols = 0 Then
               imgPiece(i).Left = imgPiece(0).Left
               imgPiece(i).Top = imgPiece(i - 1).Top + imgPiece(i - 1).Height + 50
            Else
                imgPiece(i).Left = imgPiece(i - 1).Left + imgPiece(i - 1).Width + 50
                imgPiece(i).Top = imgPiece(i - 1).Top
            End If
        End If
    Next i
    
    picPanel.Width = imgPiece(Game.Total - 1).Left + imgPiece(Game.Total - 1).Width + 250
    picPanel.Height = imgPiece(Game.Total - 1).Top + imgPiece(Game.Total - 1).Height + 250
    
    NewWidth = picPanel.Width + InitPicL + 200
    NewHeight = picPanel.Height + picPanel.Top + 800
    If NewWidth < InitSize Then
        Me.Width = InitSize
        HCenterControl picPanel
    Else
        Me.Width = NewWidth
        picPanel.Left = InitPicL
    End If
    
    If NewHeight < InitVsize Then
        Me.Height = InitVsize
    Else
        Me.Height = NewHeight
    End If
    
    
    'Me.Enabled = True

    picPanel.Visible = True

    Exit Sub
ErrHandling:
    
    Response = MsgBox("An error has occured in the StartGame() routine. Below is the number and description of the error:" & Chr(10) & Err.Number & ":" & Err.Description, vbCritical + vbAbortRetryIgnore, "Error")
    Select Case Response
        Case vbRetry
            Resume
        Case vbAbort
            Exit Sub
        Case vbIgnore
            Resume Next
    End Select

End Sub


Private Sub cmdNewGame_Click()
    StartGame
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub Form_Load()
    
    frmSplash.Show
    frmSplash.Refresh
    mciWav.FileName = App.Path & "\click.wav"
    mciWav.Command = "OPEN"
    
    'Me.BackColor = RGB(0, 0, 255)
    
    InitPicL = picPanel.Left
    InitSize = Me.Width
    InitVsize = Me.Height
    
    Randomize
    
    StartGame
    
    Wait 1
    CheckDedication

    Unload frmSplash
End Sub



Private Sub Form_Resize()
    
    Cls
    lblBestFlips.Left = Me.Width - lblBestFlips.Width - 200
    lblBestTime.Left = Me.Width - lblBestTime.Width - 200
    lnDeco.X2 = Me.Width + 200
    DrawBackGround

End Sub

Private Sub imgPiece_Click(Index As Integer)
    tmrTime.Enabled = True
    PSND
    If Game.Opened = False Then

        

        If Game.LastOpen1 = -1 Or Game.LastOpen2 = -1 Then GoTo SkipIt
        
        If imgPiece(Game.LastOpen2).Tag = imgPiece(Game.LastOpen1).Tag Then
            
            imgPiece(Game.LastOpen2).Enabled = False
            imgPiece(Game.LastOpen1).Enabled = False
            imgPiece(Game.LastOpen1).Tag = "DONE"
            imgPiece(Game.LastOpen2).Tag = "DONE"
            
        Else
        
            imgPiece(Game.LastOpen1).Picture = LoadPicture(Game.PicBack)
            imgPiece(Game.LastOpen2).Picture = LoadPicture(Game.PicBack)
            
        End If
SkipIt:
        
        Game.LastOpen1 = Index
        imgPiece(Game.LastOpen1).Enabled = False
        Game.Opened = True
        imgPiece(Index).Picture = LoadPicture(imgPiece(Index).Tag)
    Else
        Game.Moves = Game.Moves + 1
        lblFlips = "Flips:" & Game.Moves
        imgPiece(Game.LastOpen1).Enabled = True
        
        Game.LastOpen2 = Index
        Game.Opened = False
        imgPiece(Index).Picture = LoadPicture(imgPiece(Index).Tag)
        If imgPiece(Game.LastOpen2).Tag = imgPiece(Game.LastOpen1).Tag Then
            imgPiece(Game.LastOpen2).Enabled = False
            imgPiece(Game.LastOpen1).Enabled = False
            Me.picPanel.Enabled = False
            Wait 0.2
            Me.picPanel.Enabled = True
            
            imgPiece(Game.LastOpen2).Visible = False
            imgPiece(Game.LastOpen1).Visible = False
            imgPiece(Game.LastOpen1).Tag = "DONE"
            imgPiece(Game.LastOpen2).Tag = "DONE"
        End If
        
    End If
    CheckForWin
End Sub



Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpMe_Click()
    x = Shell("notepad.exe " & App.Path & "\readme.txt", 1)
End Sub

Private Sub mnuNew_Click()
    StartGame
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuRecords_Click()
    frmRecords.Show vbModal
End Sub



Private Sub tmrTime_Timer()
    Game.Time = Game.Time + 1
    lblTime = "Time:" & Game.Time & " second(s)"
End Sub
