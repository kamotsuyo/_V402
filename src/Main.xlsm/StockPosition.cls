VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iPOSITION
'#####################

Private Sub iPOSITION_Init(Order_ As Sub402.iORDER, OpenDate_ As String, OpenPrice_ As String, NeedQuantity_ As String)

End Sub
Private Sub iPOSITION_RssCloseOrder()
    Dim Result As String
    
    'äîéÆópÇ»ÇÃÇ≈RssMarginCloseOrderÇégÇ§
    'à¯êîÇ™àŸÇ»ÇÈ
'    Result = RssFOPCloseOrder_V( _
'        SerialNO, _
'        BrandCode, _
'        BuySell, _
'        OrderType, _
'        Quantity, _
'        PriceType, _
'        Price, _
'        ExecQuantityType, _
'        EXPIRYTYPE, _
'        EXPIRYDATE, _
'        OpenDate, _
'        OpenPrice, _
'        StopTriggerPrice, _
'        StopTriggerType, _
'        StopPriceType, _
        StopPrice)
    '------------------
    'ÉçÉO
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
