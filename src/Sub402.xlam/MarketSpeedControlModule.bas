Attribute VB_Name = "MarketSpeedControlModule"
Option Explicit

Const MENTENANCE_START As String = "0600"
Const MENTENANCE_END As String = "0615"
Private STARTTIME As Date
Private Ws As Object

Private myExcellName As String
Private myRaktenDir As String
Private myRaktenAPPName As String

Public Sub AutoStart()
    If Ws Is Nothing Then Set Ws = CreateObject("Wscript.shell")
    '起動時の時間帯管理
    STARTTIME = TimeValue("6:16:00")
    
    '初期値設定
    myExcellName = ThisWorkbook.Name
    myRaktenDir = "C:\Users\kamot\AppData\Local\MarketSpeed2\Bin\"
    myRaktenAPPName = "MarketSpeed2.exe"
    
    Dim mTime As String
    mTime = Format(Now(), "HHMM")
    'メンテナンス時間帯かどうかの判別
    If mTime >= MENTENANCE_START And mTime <= MENTENANCE_END Then
        FormSub.addListItem "メンテナンス時間中"
        FormSub.addListItem "開始予定時刻 " & CStr(STARTTIME)
        Application.OnTime STARTTIME, "RestartMarketSpeed"
        
    Else
        'メンテナンスではない
        Call RestartMarketSpeed
    End If
End Sub

Private Sub RestartMarketSpeed()
    If MMConf.AutoMarketSpeedFLG Then
        FormSub.addListItem ("マーケットスピード起動")
        'あらかじめMarketSpeed.exeのPATH「"C:\Users\kamot\AppData\Local\MarketSpeed2\Bin\"」を環境設定に登録してある
    
        '起動している（起動していれば）マーケットスピードのプログラムを閉じる
        Shell ("taskkill /f /im " & myRaktenAPPName)
        
        '3秒待つ
        Application.Wait Now() + TimeValue("00:00:03")
        
        'マーケットスピードを起動させる

        
        Ws.currentdirectory = myRaktenDir   'カレントディレクトリを変更
        Ws.Run myRaktenAPPName
        
        'PATHを通しているのでプログラム名だけで起動できる
        'Shell ("MarketSpeed2.exe") 'そのはずが起動エラーになるのでコメントアウト
        
        '起動を7秒待つ
        Application.Wait Now() + TimeValue("00:00:07")
        'マーケットスピードをアクティブにする（キー送出するために）
        Ws.AppActivate myRaktenAPPName  'アクティブにならない
        'ログインキーを送出する
        FormSub.addListItem ("PassWord送出")
        'ws.SendKeys "kam1824{ENTER}"
        '2024/3/20修正：マーケットスピードへのログインパスワードを変更してあるので
        Ws.SendKeys "kam1824R3Hoge{ENTER}"
        
        'MarketSpeed2起動後にエクセルをアクティブにする（キー送出するために）
        Application.Wait Now() + TimeValue("00:00:7")
        'ws.AppActivate "Primary.xlsm - Excel"
        Ws.AppActivate ThisWorkbook.Name
        
        'RSS接続開始
        '登録済みの「クイックアクセスバー」にキーを送信
        
        '時間をWAITしてもキー送出時にシートのA1に「Y」の文字が入力されてしまう
        'Application.Wait Now() + TimeValue("00:00:05")
        'ws.SendKeys "%1Y"  だめ
        'ws.SendKeys "%1{TAB}{ENTER}"  だめ
        
        '【解決<=偶然発見】以下のように疑似別スレッドで実行する
        Application.OnTime Now() + TimeValue("00:00:01"), "SendKey_MarketSpeedConnectSwitch"

        
    Else
        FormSub.addListItem ("AUTOリスタート : " & MMConf.AutoMarketSpeedFLG)
    End If
End Sub

Private Sub SendKey_MarketSpeedConnectSwitch()

    'キー内容はALT 1 ENTER
    'ws.SendKeys "%1Y"
    '2023/8/7 リボンのユーザー登録変更したら ALT 1 だけでよくなったので変更
    Ws.SendKeys "%1"        'ALT 1
    
    FormSub.addListItem ("[RSS接続]のSendKeysを実行")
End Sub

Public Sub SendKey_MarketSpeedOrderSwitch()
    '発注可にする
    If Ws Is Nothing Then Set Ws = CreateObject("Wscript.shell")
    'MarketSpeed2起動後にエクセルをアクティブにする（キー送出するために）
    'ws.AppActivate "Primary.xlsm - Excel"
    Ws.AppActivate myExcellName
    'キー内容はALT 2
    

    Ws.SendKeys "%2"
    
    FormSub.addListItem ("[発注可]のSendKey実行")
End Sub
