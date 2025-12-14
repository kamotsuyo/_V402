VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartCandle"
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
Private Index As Integer 'Chartクラスの行を格納しておく：何本前なのかの判定に使用=>TecZigzagの確定高値バックステップに使用
Private LastDateTime As Date    '最新の日時データ

Public SMA5 As Double
Public SMA25 As Double
Public EMA5 As Double
Public EMA25 As Double
'================================

'内部プロパティ
Private OpenPrice As Long
Private HighPrice As Long
Private LowPrice As Long
Private ClosePrice As Long
Private Volume As Long

Public Static Sub Init(TimeClass_ As Date, Index_ As Integer)
    TimeClass = TimeClass_
    Index = Index_
End Sub

Public Sub addData(OpenPrice_, HighPrice_, LowPrice_, ClosePrice_, Volume_)
    OpenPrice = OpenPrice_
    HighPrice = HighPrice_
    LowPrice = LowPrice_
    ClosePrice = ClosePrice_
    Volume = Volume_
End Sub


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
    iTECBASECANDLE_Price = ClosePrice  'クラスによって変わる
End Property
Private Property Get iTECBASECANDLE_Volume() As Long
    iTECBASECANDLE_Volume = Volume
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
