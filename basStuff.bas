Attribute VB_Name = "basStuff"
Sub UpdateIt()
    Dim Coll As String
    Game.Rows = GetSetting(App.Title, "Settings", "Rows", 4)
    Game.Cols = GetSetting(App.Title, "Settings", "Cols", 4)
    Coll = GetSetting(App.Title, "Settings", "Collection", "flags")
    If Dir(App.Path & "\icons\" & Coll, vbDirectory) <> "" Then
        Game.Collection = Coll
    Else
        Game.Collection = "flags"
        'MsgBox "Missing icon collection." & Chr(10) & "Selecting default(flags)...", vbExclamation, "Warning"
        SaveSetting App.Title, "Settings", "Collection", "flags"
    End If
    
    
    Game.Total = Game.Rows * Game.Cols
    Game.Record.Time = GetSetting(App.Title, "Records", Game.Total, 9999)
    Game.Record.TimeMaster = GetSetting(App.Title, "Records", Game.Total & "_NAME", "Mr. Kane")
    Game.Record.Flips = GetSetting(App.Title, "Records", Game.Total & "_FLIP", 9999)
    Game.Record.Flipper = GetSetting(App.Title, "Records", Game.Total & "_FLIP_NAME", "Mr. Kane")
End Sub
