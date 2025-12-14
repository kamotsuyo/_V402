VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DummyOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iORDER
'#####################

'規定のStanderd損益になればExitするためのイベント受信用
'ダミーの場合はFuture,Stock両方必要：(先物のダミー、株式のダミーどちらも対応するので）
Private WithEvents myFuture As Future
Attribute myFuture.VB_VarHelpID = -1
Private WithEvents myStock As Stock
Attribute myStock.VB_VarHelpID = -1

Private myAutoExit As iAUTOEXIT 'iAUTOEXIT
Private myStra As iSTRA

Private SerialNO As String
Private TimeStamp As Date
Private BuySell As String
Private Quantity As String
Private Price As String
Private ExitTimeStamp As Date
Private ProfitLoss As Long   '損益
Private isAlive As Boolean   '約定・取消済でない



'[iORDER]共通プロパティ
'=========================
Private Sub iORDER_Init(Stra_ As SubProject.iSTRA, Quantity_ As String)
    '注釈：オーダーロックは前段階のgetNewDeal内でTrueにセット済。ここでは不要
    
    '------------
    'test
    Debug.Print Stra_.DealType, Stra_.Brand.BrandName
    
    '発行時刻
    TimeStamp = Now
    'プロパティへセットする
    Set myStra = Stra_
    Quantity = Quantity_
    
    'シリアルナンバー発行
    SerialNO = NewSerialString
    
    'ディクショナリに格納
    Call DickOorder.Add(SerialNO, Me)
    
    '-----------------
    'イベント受信用
    '-----------------
    'FINANCIALPRODCTの判定
    'myFuture,myStockのセット
'    If TypeOf Stra_.Brand Is Future Then Set myFuture = Stra_.Brand
'    If TypeOf Stra_.Brand Is Stock Then Set myStock = Stra_.Brand

End Sub
Private Sub iORDER_PublishOpen(BuySell_ As SubProject.En_Buysell, Price_ As Long)

    BuySell = BuySell_
    Price = Price_
    
    'iAUTOEXITセット
    Set myAutoExit = myStra.AutoExit
    If Not myAutoExit Is Nothing Then Call myAutoExit.Init(Me)
    
    'オーダーロック解除して代わりにisAliveセット
    Call myStra.changeOrderLock(False)
    isAlive = True
    
    'F_DEAL更新
    Call F_Order.InsertDeal(Me)
    
    'フォーム更新
    Call F_Strategys.Display(myStra)
    
    '---------------------
    'DummyPosition生成実行
    '---------------------
    'Future,Stockの場合はMM_RaktenList内からCreateNewPositionがCallされる
    'DummyOrderの場合、MM_RaktenListによる更新は存在しないので、この場面で直接CreateNewPosition実行する
    Call OrderModule.CreateNewPosition(Me, CStr(TimeStamp), Price)
    
End Sub
Private Property Get iORDER_myStra() As iSTRA
    Set iORDER_myStra = myStra
End Property
Private Property Get iORDER_BuySell() As En_Buysell
    iORDER_BuySell = BuySell
End Property
Private Property Get iORDER_OrderNO() As String
    'DummyOrderには不要
End Property
Private Property Get iORDER_ProfitLoss() As Long
    iORDER_ProfitLoss = ProfitLoss
End Property
Private Property Get iORDER_isAlive() As Boolean
    iORDER_isAlive = isAlive
End Property
Private Property Get iORDER_Quantity() As Long
    iORDER_Quantity = Quantity
End Property
Private Property Get iORDER_SerialNO() As Long
    iORDER_SerialNO = SerialNO
End Property
Private Property Get iORDER_TimeStamp() As Date
    iORDER_TimeStamp = TimeStamp
End Property
Private Property Get iORDER_ExitTimeStamp() As Date
    iORDER_ExitTimeStamp = ExitTimeStamp
End Property
'=========================

Private Sub StopMe()
    Set myFuture = Nothing
    Set myStock = Nothing
    Set myAutoExit = Nothing
    
    'isAliveセット解除
    isAlive = False
End Sub

Private Sub CheckExit(iBRAND_ As iBRAND)
    '規定のStanderd損益になればExitする
    '動作確認OK:12/7
    
    Select Case BuySell
        Case En_Buysell.買
        '反対取引なので"転売"として評価する
            ProfitLoss = iBRAND_.BuyPrice - Price   '買気配値と元値の差
            Debug.Print "現在の損益", "転売", Price, iBRAND_.BuyPrice, ProfitLoss
        Case En_Buysell.売
        '反対取引なので"買戻"として評価する
            ProfitLoss = Price - iBRAND_.SellPrice
            Debug.Print "現在の損益", "買戻", Price, iBRAND_.SellPrice, ProfitLoss
    End Select
    
  
    'Exitかどうかの判定
    If ProfitLoss < -(myStra.StanderdLoss) Or ProfitLoss > myStra.StanderdMargin Then
        '------------------
        'Exitする
        '------------------
        'オーダーロック変更
        Call myStra.changeOrderLock(False)
        ExitTimeStamp = Now
        
        Call StopMe '更新を停止
        
        '--------------
        '生成元のストラテジへコールバック
        '--------------
        Call myStra.Callback(Me)
    End If


    'フォーム更新
    Call F_Order.UpdateDeal(Me)
End Sub

'------------------
'イベント受信
'------------------
Private Sub myFuture_Change(Future_ As Future)
    Call CheckExit(Future_)
End Sub

Private Sub myStock_Change(Stock_ As Stock)
    Call CheckExit(Stock_)
End Sub
