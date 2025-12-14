VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTECBASECANDLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'V302
'InterfaceにPublicは不要
'7/2　再設定
Property Get TimeClass() As Date
End Property
Property Get Index() As Integer
End Property
Property Get LastDateTime() As Date  '7/2新設
End Property
Property Get HighPrice() As Double
End Property
Property Get LowPrice() As Double
End Property
Property Get Price() As Double
End Property
Property Get Volume() As Long
End Property
Property Get SMA5() As Double
End Property
Property Get SMA25() As Double
End Property
Property Get EMA5() As Double
End Property
Property Get EMA25() As Double
End Property


