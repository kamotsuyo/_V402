VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoExitType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#####################
Implements iAUTOEXIT
'#####################
Private WithEvents myFuture As Future
Attribute myFuture.VB_VarHelpID = -1
Private WithEvents myStock As Stock
Attribute myStock.VB_VarHelpID = -1

Private myOrder As iORDER

Private Sub iAUTOEXIT_Init(Order_ As iORDER)
    
    Set myOrder = Order_
    Dim myBrand As iBRAND
    Set myBrand = myOrder.myStra.Brand
    If TypeOf myBrand Is Future Then Set myFuture = myBrand
    If TypeOf myBrand Is Stock Then Set myStock = myBrand

End Sub

Private Sub StopMe()
    '当インスタンスのイベント受信停止
    Set myFuture = Nothing
    Set myStock = Nothing
End Sub


Private Sub myFuture_Change(Future_ As Future)

End Sub

Private Sub myStock_Change(Stock_ As Stock)

End Sub
