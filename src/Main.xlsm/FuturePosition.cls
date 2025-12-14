VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FuturePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iPOSITION
'#####################

'受信するイベント
Private WithEvents MMList As MM_RaktenList
Attribute MMList.VB_VarHelpID = -1

'AutoExit用
Private myAutoExit As iAUTOEXIT

Private myStra As iSTRA
Private myOrder As iORDER '生成元

'Marketspeed2_Rss_Vba　発行に際して使用
Private SerialNO As Long             '1)発注ID [Long指定] (Initで発行する)
Private BrandCode As String          '2)銘柄コード
Private BuySell As String            '3)売買区分 : { 1: 売り , 3:買い }
Private OrderType     As String      '4)注文区分 : {0:通常注文 , 1:逆指値付通常注文 , 2:逆指値注文}  (Modifyで訂正元の注文区分を訂正する場合､{3: 通常注文⇒逆指値付通常注文､4: 逆指値付通常注文⇒通常注文}
Private Quantity As String           '5)注文数量
Private PriceType As String          '6)価格区分 : { 0:成行 , 1:指値}
Private Price As String              '7)注文価格:
Private ExecQuantityType As String   '8)執行数量条件 : {1:FOK , 2:FAK , 3:FAS}
Const EXPIRYTYPE As String = En_ExpiryType.当セッション        '9)執行時間条件 : {1:当セッション ,4:引け , 5:期間限定  , 9:取引最終日}
Const EXPIRYDATE As String = ""         '10)注文期限 : 執行時間条件が5:期間限定の場合必須。  (string YYYYMMDD)
Private StopTriggerPrice As String   '11)逆指値条件価格 : 逆指値条件価格をセット (注文区分が0:通常の場合不要)
Private StopTriggerType As String    '12)逆指値条件区分 : {1:以上 , 2:以下}
Private StopPriceType As String      '13)逆指価格区分 : {0:成行 , 1:指値}
Private StopPrice As String          '14)逆指値価格 :  逆指値価格をセット。成行の場合は省略

'Initでセット
Private OpenDate As String
Private OpenPrice As String

'IDlistイベント受信時にセットする
Private OrderNO As Double    '[Double指定]
Private OrderResult As String

'OListから取得
Private Status As String
Private StopStatus As String
Private StopCondition As String



'Rss関数結果
Private RssCloseOrderResult As String
Private RssCancelOrderResult As String
Private RssModifyOrderResult As String

'Exlistから取得
Private ExPrice As String    '約定単価
Private ExProfitLoss As String   '実現損益(Exlistの受渡金額)

'その他
Private StopMeFLG As Boolean
Private InitTimeStamp As String

Private Sub Class_Initialize()
    Set MMList = MainModule.MMList
End Sub

'-------------------------
'AutoTrailの解除
'-------------------------
Private Sub stopAutoExit()
    Set myAutoExit = Nothing

End Sub


'========================================
'iPOSITIONの共通プロパティ
Private Sub iPOSITION_Init(Order_ As Sub402.iORDER, OpenDate_ As String, OpenPrice_ As String, NeedQuantity_ As String)
    Set myOrder = Order_
    Set myStra = myOrder.myStra
    OpenDate = OpenDate_
    OpenPrice = OpenPrice_
    Quantity = NeedQuantity_
    
    'シリアルNO発行
    SerialNO = NewSerialNO
    
    '生成元オーダーのプロパティを引継ぎ
    BrandCode = myOrder.myStra.Brand.BrandCode
    Select Case myOrder.BuySell
        '反対取引にする
        Case En_Buysell.買
            BuySell = En_Buysell.売
        Case En_Buysell.売
            BuySell = En_Buysell.買
    End Select
    
    '----------------
    'ClsoeOrder生成
    '----------------
    '逆指値付通常注文で自動クローズオーダー発行
    OrderType = En_Ordertype.逆指値付通常注文
    PriceType = En_PriceType.指値
    ExecQuantityType = En_ExecQuantityType.FAS  'FAS
    StopPriceType = En_StopPriceType.指値


    '非同期処理
    Call AsyncOrderModule.SetBackPosition(Me)
    
    '-------------------
    'DicPosition
    'ディクショナリに格納
    '-------------------
    Call DicPosition.Add(SerialNO, Me)
    
End Sub

Private Sub iPOSITION_RssCloseOrder()
    Dim Result As String
    Result = RssFOPCloseOrder_V( _
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
        OpenDate, _
        OpenPrice, _
        StopTriggerPrice, _
        StopTriggerType, _
        StopPriceType, _
        StopPrice)
    '------------------
    'ログ
    If Result = "" Then
        RssCloseOrderResult = "OK"
    Else
        RssCloseOrderResult = Result
    End If
End Sub
Private Property Get iPOSITION_myStra() As Sub402.iSTRA

End Property
Private Property Get iPOSITION_OrderNO() As String

End Property
Private Property Get iPOSITION_Quantity() As String

End Property
Private Property Get iPOSITION_SerialNO() As String

End Property
Private Property Get iPOSITION_Status() As String

End Property

Private Sub MMList_IDlistChange(SerialNO_ As String, OrderNO_ As String, Result_ As String)
    If SerialNO = SerialNO_ Then
        OrderNO = OrderNO_
    
        'AutoExit起動
        Set myAutoExit = myStra.AutoExit
        If Not myAutoExit Is Nothing Then
            Call myAutoExit.Init(Me)
        End If
    End If
End Sub

Private Sub MMList_OlistChange(OrderNO_ As String, Price_ As String, Status_ As String, StopStatus_ As String, StopCondition_ As String)
    If OrderNO = OrderNO_ Then
        'ステータスなど更新
        Price = Price_  'Modifyで変更となるので更新する
        Status = Status_
        StopStatus = StopStatus_
        StopCondition = StopCondition_
    
    
        Select Case Status
            Case "約定"
                'AutoExit停止
                Call stopAutoExit
                '-------------------
                'DicPositionから削除
                '-------------------
                'ストラテジのディクショナリから削除
'                Call DicPosition.Remove(SerialNO)
                
            Case "取消済（出来無）", "取消済（出来有）"
                'AutoTrail停止
                Call stopAutoExit
                
                '-------------
                '成行でExitを自動発行
                '-------------
                'クロースオーダーキャンセル時は強制的に成行EXITのクローズオーダー発行を行う
                '新規クローズオーダーを生成する
'                Dim newCloseOrder As CloseOrder
'                Set newCloseOrder = New CloseOrder
'                Call newCloseOrder.Init(Me.myOpenOrder, Me.OpenDate, Me.OpenPrice)
'
'                Call newCloseOrder.PublishCloseMarketExit
                
                'このクローズオーダーをディクショナリから削除
'                Call DicPosition.Remove(Me.SerialNO)
        End Select
    End If
End Sub
'========================================
