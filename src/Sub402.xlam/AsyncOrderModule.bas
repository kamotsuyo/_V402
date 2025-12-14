Attribute VB_Name = "AsyncOrderModule"
Option Explicit

'---------------------
'非同期処理
'CallByNameを使用する

Private CollStack As Collection


Public Sub SetBackPosition(Position_ As iPOSITION)
    If CollStack Is Nothing Then Set CollStack = New Collection
    
    'コレクションに格納
    Call CollStack.Add(Position_)
    
    If CollStack.Count > 0 Then
        '-------------
        'Application.OnTimeで非同期処理
        '-------------
        Application.OnTime Now, "AsyncOrderModule.ExecAsync"
    
    End If
End Sub


Public Sub ExecAsync()

    'CollStackをループ処理
    Do While CollStack.Count > 0
        Dim myPosition As iPOSITION
        Set myPosition = CollStack.Item(1) '一番古いアイテム：コレクションは１オリジン)
        
        'RssCloseOrderをCallして実行する
        Call myPosition.RssCloseOrder
    
        Call CollStack.Remove(1)
    Loop

End Sub


