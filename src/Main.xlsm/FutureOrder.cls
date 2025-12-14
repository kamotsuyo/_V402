VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FutureOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iORDER
'#####################

'受信するイベント
Private WithEvents MMList As MM_RaktenList
Attribute MMList.VB_VarHelpID = -1
Private WithEvents myFuture As Future
Attribute myFuture.VB_VarHelpID = -1
Private myiAUTOEXIT As iAUTOEXIT 'iAUTOEXIT

'Class_Initializeでセットするプロパティ
Private TimeStamp As Date

'iORDER_Initでセットするプロパティ
Private myStra As iSTRA
'SerialNO,BrandCodeもDEAL_Initでセット

Private ExitTimeStamp As Date
Private isAlive As Boolean
Private ProfitLoss As Long

'***********Marketspeed2_Rss_Vba発行に際して使用するプロパティ***********
'iORDER_Initでセットするプロパティ
Private SerialNO As Long             '1)発注ID [Long指定] (Initで発行する)
Private BrandCode As String          '2)銘柄コード
Private Quantity As String           '5)注文数量
'iORDER_Publishでセットするプロパティ
Private BuySell As String            '3)売買区分 : { 1: 売り , 3:買い }
Private Price As String              '7)注文価格:
Private OrderType As String          '4)注文区分 : {0:通常注文 , 1:逆指値付通常注文 , 2:逆指値注文}  (Modifyで訂正元の注文区分を訂正する場合､{3: 通常注文⇒逆指値付通常注文､4: 逆指値付通常注文⇒通常注文}
Private PriceType As String          '6)価格区分 : { 0:成行 , 1:指値}
Private ExecQuantityType As String   '8)執行数量条件 : {1:FOK , 2:FAK , 3:FAS}

'以下は注文区分が0:通常の場合 初期値""でよい
Private StopTriggerPrice As String   '11)逆指値条件価格 : 逆指値条件価格をセット (注文区分が0:通常の場合不要)
Private StopTriggerType As String    '12)逆指値条件区分 : {1:以上 , 2:以下}
Private StopPriceType As String      '13)逆指価格区分 : {0:成行 , 1:指値}
Private StopPrice As String          '14)逆指値価格 :  逆指値価格をセット。成行の場合は省略

'固定でよい
Const EXPIRYTYPE = En_ExpiryType.当セッション '9)執行時間条件 : {1:当セッション ,4:引け , 5:期間限定  , 9:取引最終日}
Const EXPIRYDATE As String = ""        '10)注文期限 : 執行時間条件が5:期間限定の場合必須。  (string YYYYMMDD)
'********************************************

'IDlistから取得
Private OrderResult As String 'Rss関数結果
Private OrderNO As String
'Olistから取得
Private Status As String
Private StopStatus As String
Private StopCondition As String


Private Sub Class_Initialize()
    Set MMList = MainModule.MMList
    TimeStamp = Now
End Sub


'[iORDER]共通プロパティ
'=========================
Private Sub iORDER_Init(Stra_ As Sub402.iSTRA, Quantity_ As String)
    '注釈：オーダーロックは前段階のgetNewDeal内でTrueにセット済。ここでは不要
    '--------------
    'シリアルナンバー発行する
    '--------------
    SerialNO = NewSerialString
    
    '-----------------
    '引数からOpen用プロパティセットする
    '-----------------
    Set myStra = Stra_
    BrandCode = Stra_.Brand.BrandCode
    Quantity = Quantity_
    
    '-----------------
    'ディクショナリに格納
    '-----------------
    Call DickOorder.Add(SerialNO, Me)
    
    '-----------------
    'イベント受信用インスタンス生成
    '-----------------
    'このFutureOrderクラスが生成されるということは取扱商品Brandは先物である：myFutureのセット
    Set myFuture = Stra_.Brand
    
    
    'オーダーロック解除して代わりにisAliveセット
    Call myStra.changeOrderLock(False)
    isAlive = True
    
End Sub
Private Sub iORDER_PublishOpen(BuySell_ As Sub402.En_Buysell, Price_ As Long)

    BuySell = BuySell_
    Price = Price_
    OrderType = En_Ordertype.通常注文
    PriceType = En_PriceType.指値
    ExecQuantityType = En_ExecQuantityType.FAS
    
    '----------------
    'FOPOptneOrder発行
    '----------------
    '直接Rss関数を発行する
    Call RssFOPOpenOrder
    
    '----------------
    'フォーム表示更新
    '----------------
    'F_DEAL更新
    Call F_Order.InsertDeal(Me)
    'フォーム更新
    Call F_Strategys.Display(myStra)
End Sub
Private Property Get iORDER_myStra() As iSTRA
    Set iORDER_myStra = myStra
End Property
Private Property Get iORDER_OrderNO() As String
    iORDER_OrderNO = OrderNO
End Property
Private Property Get iORDER_BuySell() As Sub402.En_Buysell
    iORDER_BuySell = BuySell
End Property
Private Property Get iORDER_Quantity() As Long
    iORDER_Quantity = Quantity
End Property
Private Property Get iORDER_SerialNO() As Long
    iORDER_SerialNO = SerialNO
End Property
Private Property Get iORDER_ProfitLoss() As Long
    iORDER_ProfitLoss = ProfitLoss
End Property
Private Property Get iORDER_isAlive() As Boolean
    iORDER_isAlive = isAlive
End Property
Private Property Get iORDER_TimeStamp() As Date
    iORDER_TimeStamp = TimeStamp
End Property
Private Property Get iORDER_ExitTimeStamp() As Date
    iORDER_ExitTimeStamp = ExitTimeStamp
End Property

'=========================
'Rssオーダーメソッド
Private Sub RssFOPOpenOrder()

    Dim Result As String
    Result = Marketspeed2_Rss_Vba.RssFOPOpenOrder_V( _
        SerialNO, _
        BrandCode, _
        BuySell, _
        OrderType, _
        Quantity, _
        PriceType, _
        Price, _
        ExecQuantityType, _
        EXPIRYTYPE, _
        EXPIRYDATE, _
        StopTriggerPrice, _
        StopTriggerType, _
        StopPriceType, _
        StopPrice _
        )
    
    '------------------------
    '戻り値Resultについて
    '正常終了ならば""空白になる
    Select Case Result
        Case ""
            '正常終了
            OrderResult = "正常終了" 'このクラスのプロパティにセットする
        Case Else
            OrderResult = Result
            'エラーあり：MarketSpeedの[発注不可]になっているなどの要因でRssFOPOpenOrder_Vメドッドの戻り値が""出ない場合の処理
            
    End Select
    
    Debug.Print OrderResult
End Sub

'-------------
'イベント受信用
'-------------
Private Sub MMList_IDlistChange(SerialNO_ As String, OrderNO_ As String, Result_ As String)
    '-------------------
    'OrderNOを格納する
    '-------------------
    If SerialNO = SerialNO_ Then
        'オーダーNOをセット
        OrderNO = OrderNO_
        OrderResult = Result_
    
        If Result_ <> "正常終了" And Result_ <> "発注済み" Then
            Call DickOorder.Remove(SerialNO_) 'RSS結果エラーにつき、このORDERのインスタンスをディクショナリから削除する
        End If
    End If

End Sub

Private Sub MMList_OlistChange(OrderNO_ As String, Price_ As String, Status_ As String, StopStatus_ As String, StopCondition_ As String)
    If OrderNO = OrderNO_ Then
        Price = Price_  'Modifyで変更となるので更新する
        Status = Status_
        StopStatus = StopStatus_
        StopCondition = StopCondition_
        
        'isAliveの設定
        isAlive = isAliveStatus(Status_)
        
        Debug.Print "MMList_OlistChange"; Price, Status, StopStatus, StopCondition
        
        '-------------
        '自動クローズオーダー発行
        '-------------
        If Status = "約定" Then
            Call MMList.SetBackOrder(Me)
        End If
        
    End If
End Sub

Private Sub myFuture_Change(Future_ As Future)

End Sub
