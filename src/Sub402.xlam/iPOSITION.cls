VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iPOSITION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'************************
'Sub iPosition
'V303
'************************
'このオブジェクトのプロパティ設定：Instancing は [2-PublicNotCreatable]にしておく

Public Sub Init(Order_ As iORDER, OpenDate_ As String, OpenPrice_ As String, NeedQuantity_ As String)
End Sub
Public Sub RssCloseOrder()
End Sub
Public Sub PublishAutoClose()
End Sub
Property Get SerialNO() As String
End Property
Property Get OrderNO() As String
End Property
Property Get Status() As String
End Property
Property Get Quantity() As String
End Property
Property Get myStra() As iSTRA
End Property
