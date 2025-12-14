Attribute VB_Name = "AsyncModule"
Option Explicit

'非同期処理のためのモジュール
Private CollStack As Collection
Private DicDelayStack As Dictionary

Public Sub setStack(Obj_ As iASYNCGATE, Optional Args_ As Variant = Empty)
    On Error GoTo ErrHandler
    If CollStack Is Nothing Then Set CollStack = New Collection
    
    Dim Arr As Variant
    Arr = Array(Obj_, Args_)
    
    'コレクションに格納
    Call CollStack.Add(Arr)
    
    If CollStack.Count > 0 Then
        '-------------
        'Application.OnTimeで非同期処理
        '-------------
        Application.OnTime Now, "AsyncModule.execAsync"
    
    End If
Exit Sub
ErrHandler:
    Err.Source = "AsyncOrderModule.setCloseOrder"   '発生元がわかるように命名する
    '--------以下は共通--------
    Call F_Error.DisplayError(Err)

End Sub

'------------
'Delay用
Public Sub setDelayStack(Obj_ As iASYNCGATE, AddTime As Date)
    
    If DicDelayStack Is Nothing Then Set DicDelayStack = New Dictionary
    
    Static KeyNO As Integer
    KeyNO = KeyNO + 1
    
    Call DicDelayStack.Add(CStr(KeyNO), Obj_)
    
    If DicDelayStack.Count > 0 Then
        Dim ProcString As String
        ProcString = getOntimeProcedureString("AsyncModule.execDelayAsync", Array(KeyNO))
        
        '-------------------
        'Application.OnTime関数を使用
        '-------------------
        Application.OnTime Now + AddTime, ProcString
    
    End If
End Sub


Public Sub ExecAsync()
    If CollStack Is Nothing Then Exit Sub
    Do While CollStack.Count > 0

        Dim Arr As Variant
        Arr = CollStack.Item(1) '一番古いアイテム：コレクションは１オリジン)
        
        Dim Obj As iASYNCGATE
        Set Obj = Arr(0)
        Dim Args As Variant
        Args = Arr(1)
        
        Call Obj.GATE(Args)

        Call CollStack.Remove(1)
    Loop
End Sub


Public Sub execDelayAsync(Key_ As String)
    
    Dim Obj As iASYNCGATE
    If Not DicDelayStack.Item(Key_) Is Nothing Then
        '11/1停止エラー発生のため、上記回避策を追加
        Set Obj = DicDelayStack.Item(Key_)
        
        Call Obj.GATE(Empty)
        
        Call DicDelayStack.Remove(Key_)
    End If
End Sub

'補助関数
Public Function getOntimeProcedureString(FuncName_ As String, Args_ As Variant) As String
    getOntimeProcedureString = " '" & FuncName_ & " " & """" & Join(Args_, """,""") & """" & " '"
End Function
