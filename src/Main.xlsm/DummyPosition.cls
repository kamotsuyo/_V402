VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DummyPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iPOSITION
'#####################

Private WithEvents myFuture As Future
Attribute myFuture.VB_VarHelpID = -1
Private WithEvents myStock As Stock
Attribute myStock.VB_VarHelpID = -1

Private myAutoExit As iAUTOEXIT 'iAUTOEXIT
Private myOrder As iORDER
Private OpenPrice As Long

Private Sub iPOSITION_Init(Order_ As SubProject.iORDER, OpenDate_ As String, OpenPrice_ As String)
    Set myOrder = Order_
    OpenPrice = OpenPrice_
    
    '-----------------
    'イベント受信開始
    '-----------------
    Dim myBrand As iBRAND
    Set myBrand = Order_.myStra.Brand
    'FINANCIALPRODCTの判定
    If TypeOf myBrand Is Future Then Set myFuture = myBrand
    If TypeOf myBrand Is Stock Then Set myStock = myBrand
    
    Debug.Print "iPOSITION_Init", myBrand.BrandName, "HOgehoge"
End Sub
Private Property Get iPOSITION_myStra() As SubProject.iSTRA

End Property
Private Property Get iPOSITION_OrderNO() As String

End Property
Private Property Get iPOSITION_Quantity() As String

End Property
Private Property Get iPOSITION_SerialNO() As String

End Property
Private Property Get iPOSITION_Status() As String

End Property

Private Sub StopMe()
    Set myFuture = Nothing
    Set myStock = Nothing
    Set myAutoExit = Nothing
    
    'isAliveセット解除
'    isAlive = False
End Sub

Private Sub CheckExit(iBRAND_ As iBRAND)
    '規定のStanderd損益になればExitする
    '動作確認OK:12/7
    
    Select Case myOrder.BuySell
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
'        Call myStra.Callback(Me)
    End If


    'フォーム更新
'    Call F_Order.UpdateDeal(Me)
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
