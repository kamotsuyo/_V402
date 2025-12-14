VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'********************
'Instancing := 2 - publicNotCreatableにしておく
'********************

Private WithEvents mySheet As Worksheet
Attribute mySheet.VB_VarHelpID = -1

Private myForm As FormSub

'接続完了済みかどうかのフラグ
Private ConnectedFLG As Boolean

'発行するイベント
Event RssConnectChange(Status_ As String)
Event ReConnect()

Private Sub Class_Initialize()
    '記述済のワークシートをインスタンスセット
'    Set mySheet = ThisWorkbook.Sheets("RssConn")
    Set mySheet = getMySheet("Conn", ThisWorkbook)
    With mySheet
        .Range("a1").Formula = "=RssFOPCapacityList()"
'        .Visible = xlSheetHidden '非表示
    End With
    
    'フォーム
    Set myForm = FormSub
    Call myForm.addListItem("Conn起動：更新待ち")
End Sub

Private Sub mySheet_Change(ByVal Target As Range)
    '同じステータスなのにイベント発行を防止する必要がある
    If Target.Address <> "$A$1" Then Exit Sub
    
    Static PreviousStatus As String
    
    Dim TragetString As String
    TragetString = mySheet.Range("a1").Value

    Dim Status As String
    '=>のあとの文字列を抽出しstatusに格納する
    Const Drill As String = "=> "
    Status = Mid(TragetString, InStr(TragetString, Drill) + Len(Drill))
    
    If Status = PreviousStatus Then
        Exit Sub
    Else
        '前回ステータスを更新しておく
        PreviousStatus = Status
    End If
    
    '*************************
    '接続状態判別開始
    Select Case Status
        Case "応答待ち"
            myForm.Image1.BackColor = vbYellow
            
        Case "接続待ち"
            'ステータス変更
            myForm.addListItem ("RSS接続状態 = " & Status)
            myForm.Image1.BackColor = vbRed

            '=========================
            '接続【待ち・切断】
            RaiseEvent RssConnectChange(Status)
            '=========================

        Case "完了"
            'ステータス変更
            myForm.addListItem ("RSS接続状態 = " & Status)
            myForm.Image1.BackColor = vbGreen
           
    
            If ConnectedFLG = False Then
                ConnectedFLG = True
            Else
                '=========================
                RaiseEvent ReConnect
                '=========================
            End If
            
            '=========================
            '接続【完了】
            RaiseEvent RssConnectChange(Status)
            '=========================
            
            '**************************
            'SendKey実行
            'MarketSpeed発注不可==>発注可切り替え
    '        Call ControlModule.SendKey_MarketSpeedOrderSwitch
            'V101:2024/12/9
            'ときおり、sendKeyの効果がなく[発注不可]のままになるため、Ontimeで実行するように変更。
            '**************************
            
            '2025/12/6 6:15初回起動時にエラーで停止。
            '対策として、下記の待機時間を2秒に延長する:FutureCodeResolverの待機時間と同じだが問題ないはず。
'            Application.OnTime Now + TimeValue("00:00:01"), "SendKey_MarketSpeedOrderSwitch"
'            Application.OnTime Now + TimeValue("00:00:02"), "SendKey_MarketSpeedOrderSwitch"
        
        Case Else
            Debug.Print "Else", Status
        
    End Select
End Sub

