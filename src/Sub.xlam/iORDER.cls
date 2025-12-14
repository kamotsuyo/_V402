VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**********************
'DummyOrderとCloseOrderの親クラス

'InterfaceにPublic は不要
'AutoExitにおいてmyStraを取得する
Sub Init(Stra_ As iSTRA, Quantity_ As String)   'ストラテジと数量を引数にする
End Sub
Sub PublishOpen(BuySell_ As En_Buysell, Price_ As Long)
End Sub
Property Get myStra() As iSTRA
End Property
Property Get OrderNO() As String
End Property
Property Get BuySell() As En_Buysell   '    売 = 1    買 = 3
End Property
Property Get ProfitLoss() As Long
End Property
Property Get isAlive() As Boolean   'Exitや取消済でない
End Property
Property Get Quantity() As Long
End Property
Property Get SerialNO() As Long
End Property
Property Get TimeStamp() As Date
End Property
Property Get ExitTimeStamp() As Date
End Property

