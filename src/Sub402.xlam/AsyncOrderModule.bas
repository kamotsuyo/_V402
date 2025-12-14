Attribute VB_Name = "AsyncOrderModule"
Option Explicit

'---------------------
'非同期処理
'CallByNameを使用する

Private CollStack As Collection


Public Sub SetCloseOrder(Order_ As iORDER)
    If CollStack Is Nothing Then Set CollStack = New Collection
    
    'コレクションに格納
    Call CollStack.Add(Order_)
    
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
        Dim myOrder As iORDER
        Set myOrder = CollStack.Item(1) '一番古いアイテム：コレクションは１オリジン)
        
        Call myOrder.CallbackAsyncCloseOrder
    
        Call CollStack.Remove(1)
    Loop

End Sub


