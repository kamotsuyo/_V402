VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iBRAND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************
'FINANCIALPRODUCT
'Future,Stockの親クラス
'*******************


Property Get BrandCode() As String
End Property
Property Get PetName() As String
End Property
Property Get BrandName() As String
End Property
Property Get Price() As Long
End Property
Property Get TimeStamp() As String
End Property
Property Get SellPrice() As Long '売気配
End Property
Property Get SellQuantity() As Long
End Property
Property Get BuyPrice() As Long '買気配
End Property
Property Get BuyQuantity() As Long
End Property
Property Get PreviousPrice() As Long
End Property
Property Get PreviousRatio() As Double
End Property
Property Get SellBuyRatio() As Double
End Property
Property Get TheoricPrice() As Double
End Property

