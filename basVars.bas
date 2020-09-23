Attribute VB_Name = "basVars"
Global Game As GameType
Global FSO As New FileSystemObject
Global InitSize As Integer
Global InitVsize As Integer
Global InitPicL As Integer
Type RecordType
    Time As Integer
    Flips As Integer
    TimeMaster As String
    Flipper As String
End Type
Type GameType
    Cols As Integer
    Rows As Integer
    Collection As String
    Total As String
    PicBack As String
    Moves As Integer
    Time As Integer
    Record As RecordType
    Opened As Boolean
    LastOpen1 As Integer
    LastOpen2 As Integer
End Type


