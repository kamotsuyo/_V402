VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TickCandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iTECBASECANDLE
'#####################

'================================
'iTECBASE共通プロパティ
Private TimeClass As Date
Private Index As Integer
Private LastDateTime As Date    '最新の日時データ

'Tickクラスにてセットするプロパティ
Public SMA5 As Double
Public SMA25 As Double
Public EMA5 As Double
Public EMA25 As Double
'================================

'内部プロパティ
Private TotalVolume As Long   '合計出来高
Private TotalAmount As LongLong  '合計取引金額:LongLongに変更
Private HighPrice As Double
Private LowPrice As Double


'5/17追記:TickCandleのみ
Public AddedCount As Integer    'Tickから参照あり

Public Sub Init(TimeClass_ As Date, Index_ As Integer)
    TimeClass = TimeClass_
    Index = Index_

End Sub

Public Sub addData(CurrentTime_ As Date, Volume_ As Long, Price_ As Long)

    '4/11 プロパティ抽出時に適用する最低対象出来高
    Const MINIMUMVOLUME As Integer = 1

    '5/17追記
    'AddedCountの計算処理
    If CurrentTime_ = LastDateTime Then
        Me.AddedCount = Me.AddedCount + 1
    Else
        Me.AddedCount = 1   '初期化
        LastDateTime = CurrentTime_
    End If
    
    '---------------------
    '自身のプロパティに格納
    '最低対象出来高を適用する
    '4/11追記
    '---------------------
    If Volume_ >= MINIMUMVOLUME Then
        'TotalVolume,TotalAmount
        TotalVolume = TotalVolume + Volume_
        TotalAmount = TotalAmount + Volume_ * Price_
        'HighPrice,LowPrice
        If HighPrice = 0 Or Price_ > HighPrice Then HighPrice = Price_
        If LowPrice = 0 Or Price_ < LowPrice Then LowPrice = Price_
    End If

End Sub

'現在時間の階級時間からの経過秒数位置
'7/3現在、使用していないが残しておく
Public Function getElapsedIndex() As Integer
    getElapsedIndex = DateDiff("S", TimeClass, LastDateTime) + 1
End Function

'=================================
'iTECBASECANDLE共通プロパティ
Private Property Get iTECBASECANDLE_Index() As Integer
    iTECBASECANDLE_Index = Index
End Property
Private Property Get iTECBASECANDLE_TimeClass() As Date
    iTECBASECANDLE_TimeClass = TimeClass
End Property
Private Property Get iTECBASECANDLE_LastDateTime() As Date
    iTECBASECANDLE_LastDateTime = LastDateTime
End Property

Private Property Get iTECBASECANDLE_HighPrice() As Double
    iTECBASECANDLE_HighPrice = HighPrice
End Property
Private Property Get iTECBASECANDLE_LowPrice() As Double
    iTECBASECANDLE_LowPrice = LowPrice
End Property
Private Property Get iTECBASECANDLE_Price() As Double
'    iTECBASECANDLE_Price = Me.VWAP 'クラスによって変わる
    iTECBASECANDLE_Price = TotalAmount / TotalVolume 'クラスによって変わる
End Property
Private Property Get iTECBASECANDLE_Volume() As Long
    iTECBASECANDLE_Volume = TotalVolume    '合計出来高
End Property
Private Property Get iTECBASECANDLE_SMA5() As Double
    iTECBASECANDLE_SMA5 = SMA5
End Property
Private Property Get iTECBASECANDLE_SMA25() As Double
    iTECBASECANDLE_SMA25 = SMA25
End Property
Private Property Get iTECBASECANDLE_EMA5() As Double
    iTECBASECANDLE_EMA5 = EMA5
End Property
Private Property Get iTECBASECANDLE_EMA25() As Double
    iTECBASECANDLE_EMA25 = EMA25
End Property
'=================================
