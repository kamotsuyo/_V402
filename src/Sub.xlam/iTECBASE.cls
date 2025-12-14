VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTECBASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'V302
'InterfaceにPublicは不要

'***注意点***
'Event Change(Tec_ as iTECBASE) といったイベントを作成することはできない
'==> エラー459になる
'************

Property Get Candle(Optional I_ As Integer = 0) As iTECBASECANDLE
'引数I_は計測開始起点：0なら最新足。1なら1本前の足を起点とする
End Property
Property Get MostHighCandle(Optional N_ As Integer = 5, Optional I_ As Integer = 0) As iTECBASECANDLE
'引数N_は計測対象期間：5なら最新足から5本前までを集計する
'引数I_は計測開始起点：0なら最新足。1なら1本前の足を起点とする
End Property
Property Get MostLowCandle(Optional N_ As Integer = 5, Optional I_ As Integer = 0) As iTECBASECANDLE
End Property
Property Get Count() As Integer
End Property
'7/2
Property Get PriceRCI(Optional N_ As Integer = 10) As Double
End Property
Property Get PriceRSI(Optional N_ As Integer = 10) As Double
End Property
Property Get EMA5() As Double
End Property
Property Get EMA25() As Double
End Property
Property Get VolumeRSI(Optional N_ As Integer = 5) As Double
End Property
Property Get EMA5RSI(Optional N_ As Integer = 5) As Double
End Property
