Attribute VB_Name = "basCoolTools"
'
'CoolTools .BAS file for VB
'from ALBULESCU 'S
'v1.0
'


Global MoveX(300) As Integer
Global MoveY(300) As Integer
Global MoveNumber(300) As Integer
Sub Wait(Interval As Double)
EndTime = Timer + Interval
While EndTime > Timer: DoEvents: Wend

End Sub



Sub CoolPrintFF(File As String, ControlOrForm, Optional Interval As Double)
If Interval = 0 Then Interval = 0.04

FF = FreeFile
Open File For Input As FF
While Not EOF(FF)
s = Input(1, FF)
If s <> Chr(13) Then ControlOrForm.Print s;
T = Timer + Interval: While T > Timer: DoEvents: Wend
Wend
Close FF

End Sub


Sub CoolPrint(StringToPrint As String, ControlOrForm, Optional Interval As Double)
If Interval = 0 Then Interval = 0.04
Open "C:\temp.dat" For Output As #1
Print #1, StringToPrint
Close #1
FF = FreeFile
Open "C:\temp.dat" For Input As FF
While Not EOF(FF)
s = Input(1, FF)
If s <> Chr(13) Then ControlOrForm.Print s;
T = Timer + Interval: While T > Timer: DoEvents: Wend
Wend
Close FF
Kill "C:\temp.dat"
End Sub
Sub CoolSLC(StringToPrint As String, LBL As Label, Optional Interval As Double)
LBL.Alignment = 0
LBL.AutoSize = True
LBL.Caption = ""

If Interval = 0 Then Interval = 0.04
Open "C:\temp.dat" For Output As #1
Print #1, StringToPrint
Close #1
FF = FreeFile
Open "C:\temp.dat" For Input As FF
While Not EOF(FF)
s = Input(1, FF)
If s <> Chr(13) Then LBL.Caption = LBL.Caption & s
T = Timer + Interval: While T > Timer: DoEvents: Wend
Wend
Close FF
Kill "C:\temp.dat"
End Sub
Sub CoolCLC(LBL As Label, Optional Interval As Double)
If Interval = 0 Then Interval = 0.04
LBL.AutoSize = True
LBL.Alignment = 1
Do Until LBL = ""
LBL.Caption = Right(LBL, Len(LBL.Caption) - 1)
T = Timer + Interval: While T > Timer: DoEvents: Wend
Loop

End Sub


Sub CenterWindow(Form As Form)
Form.Left = Screen.Width / 2 - Form.Width / 2
Form.Top = Screen.Height / 2 - Form.Height / 2
End Sub

Sub CenterControl(control As control)
control.Left = control.Container.Width / 2 - control.Width / 2
control.Top = control.Container.Height / 2 - control.Height / 2

End Sub

Sub HCenterControl(control As control)
control.Left = control.Container.Width / 2 - control.Width / 2


End Sub


Sub SavePos(Form As Form, Key As String)

On Error Resume Next


If Form.WindowState = 0 Then

SaveSetting App.Title, "Position", "Height(" & Key & ")", Form.Height
SaveSetting App.Title, "Position", "Width(" & Key & ")", Form.Width
SaveSetting App.Title, "Position", "Top(" & Key & ")", Form.Top
SaveSetting App.Title, "Position", "Left(" & Key & ")", Form.Left
SaveSetting App.Title, "Position", "WindowState(" & Key & ")", Form.WindowState
Else

oldWindowState = Form.WindowState
Form.WindowState = 0

SaveSetting App.Title, "Position", "Height(" & Key & ")", Form.Height
SaveSetting App.Title, "Position", "Width(" & Key & ")", Form.Width
SaveSetting App.Title, "Position", "Top(" & Key & ")", Form.Top
SaveSetting App.Title, "Position", "Left(" & Key & ")", Form.Left
SaveSetting App.Title, "Position", "WindowState(" & Key & ")", oldWindowState

Form.WindowState = oldWindowState
End If


End Sub

Sub LoadPos(Form As Form, Key As String, WndResize As Boolean, WndState As Boolean)
On Error Resume Next



Form.Top = GetSetting(App.Title, "Position", "Top(" & Key & ")", Form.Top)
Form.Left = GetSetting(App.Title, "Position", "Left(" & Key & ")", Form.Left)

If WndResize = True Then
Form.Height = GetSetting(App.Title, "Position", "Height(" & Key & ")", Form.Height)
Form.Width = GetSetting(App.Title, "Position", "Width(" & Key & ")", Form.Width)
End If

If WndState = True Then
formWindowState = GetSetting(App.Title, "Position", "WindowState(" & Key & ")", Form.WindowState)
If formWindowState = 2 And Form.MaxButton = False Then
Exit Sub
End If
If formWindowState = 1 And Form.MinButton = False Then
Exit Sub
End If
Form.WindowState = GetSetting(App.Title, "Position", "WindowState(" & Key & ")", Form.WindowState)
End If

End Sub

Sub MoveSprite(control As control, Key As Integer, Optional XMoveSize As Integer, Optional YMoveSize, Optional RandomizeIfUnKknown As Boolean, Optional RandomizeSometimes As Boolean, Optional MovesBeforeRandomize As Integer, Optional Minus As Integer, Optional Plus As Integer)
On Error Resume Next

MoveNumber(Key) = MoveNumber(Key) + 1

If RandomizeIfUnKknown = True Then
tmpNumberOne = IIf(Rnd * 1 > 0.6, 2, 1)
tmpNumberTwo = IIf(Rnd * 1 > 0.6, 2, 1)

If MoveY(Key) = 0 Then MoveY(Key) = tmpNumberOne
If MoveX(Key) = 0 Then MoveX(Key) = tmpNumberTwo
End If

If RandomizeSometimes = True Then
If MoveNumber(Key) = MovesBeforeRandomize Then
MoveNumber(Key) = 0
MoveY(Key) = tmpNumberOne
MoveX(Key) = tmpNumberTwo
End If
End If



If control.Top < 0 + Plus Then
MoveY(Key) = 2
End If

If control.Top > control.Container.Height - control.Height - Minus Then
MoveY(Key) = 1
End If

If control.Left < 0 + Plus Then
MoveX(Key) = 2
End If

If control.Left > control.Container.Width - control.Width - Minus Then
MoveX(Key) = 1
End If



If MoveY(Key) = 1 Then
control.Top = control.Top - YMoveSize
Else
control.Top = control.Top + YMoveSize
End If



If MoveX(Key) = 1 Then
control.Left = control.Left - XMoveSize
Else
control.Left = control.Left + XMoveSize
End If


End Sub


