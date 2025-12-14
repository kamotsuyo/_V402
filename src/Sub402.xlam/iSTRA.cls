VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSTRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Ver 302

Option Explicit

'InterfaceにPublicは不要
Sub Init(Brand_ As iBRAND, DealType_ As En_DealType)
End Sub
Sub changeOrderLock(Flg_ As Boolean)
End Sub
Sub Callback(Order_ As iORDER)
End Sub
Property Get isOutSideHours() As Boolean '対象時間外かどうか
End Property
Property Get DealType() As En_DealType
End Property
Property Get AliveCount() As Long
End Property
Property Get TotalCount() As Long
End Property
Property Get STANDERDMARGIN() As Long
End Property
Property Get STANDERDLOSS() As Long
End Property
Property Get Name() As String
End Property
Property Get Brand() As iBRAND   'DummyOrderで使用中
End Property
Property Get OrderLock() As Boolean
End Property
Property Get CurrentProfitLoss() As Long
End Property
Property Get TotalProfitLoss() As Long
End Property
Property Get AutoExit() As iAUTOEXIT
End Property
Property Get MAXORDERQUANTITY() As Integer '最大発注可能件数
End Property
