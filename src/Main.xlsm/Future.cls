VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Future"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iBRAND
'#####################

'*************************
'このMarketクラス内部で先物と株式の判定と処理分岐を行う
'先物ならRssFOPMarket
'株式ならRssMarket
'を使用する
'*************************

Private WithEvents mySheet As Worksheet
Attribute mySheet.VB_VarHelpID = -1

'以下、すべてPrivate :iBRAND公開プロパティで対応するので
'初期Initにてセットする
Private BrandCode As String
Private PetName As String

'mySheet_Calculateイベント受信時にセットする
Private BrandName As String
Private Price As Long
Private TimeStamp As String
Private SellPrice As Long '売気配
Private SellQuantity As Long
Private BuyPrice As Long '買気配
Private BuyQuantity As Long
Private PreviousPrice As Long
Private PreviousRatio As Double

'発行するイベント
Event Change(Future_ As Future)

'処理時間計測
Private El As ElapsedTime

Private Sub Class_Initialize()
    '処理時間計測
    Set El = New ElapsedTime
End Sub

Public Sub Init(FutureCode_ As FutureCode)

    'シートの銘柄項目を作る配列
    Dim ArrKoumoku As Variant
    ArrKoumoku = Array( _
        "銘柄名称", _
        "現在値", _
        "現在値詳細時刻", _
        "最良売気配値1", _
        "最良売気配数量1", _
        "最良買気配値1", _
        "最良買気配数量1", _
        "前日終値", _
        "前日比率")
        
    
    'マーケット情報取得部分をワークシートに記述する

    
    BrandCode = FutureCode_.BrandCode
    PetName = FutureCode_.PetName
    
    Set mySheet = getMySheet("M_" & BrandCode)
    
    Dim i As Integer
    For i = LBound(ArrKoumoku) To UBound(ArrKoumoku)
        mySheet.Cells(i + 1, 1).Value = ArrKoumoku(i)
        'RssFOPMarketを使う
        mySheet.Cells(i + 1, 2).Formula = "=RssFOPMarket(""" & BrandCode & """,""" & ArrKoumoku(i) & """)"
    Next i

    
End Sub

'===================================
'iBRAND共通プロパティ
Private Property Get iBRAND_BrandCode() As String
    iBRAND_BrandCode = BrandCode
End Property
Private Property Get iBRAND_PetName() As String
    iBRAND_PetName = PetName
End Property
Private Property Get iBRAND_BrandName() As String
    iBRAND_BrandName = BrandName
End Property
Private Property Get iBRAND_Price() As Long
    iBRAND_Price = Price
End Property
Private Property Get iBRAND_TimeStamp() As String
    iBRAND_TimeStamp = TimeStamp
End Property
Private Property Get iBRAND_SellPrice() As Long
    iBRAND_SellPrice = SellPrice
End Property
Private Property Get iBRAND_SellQuantity() As Long
    iBRAND_SellQuantity = SellQuantity
End Property
Private Property Get iBRAND_BuyPrice() As Long
    iBRAND_BuyPrice = BuyPrice
End Property
Private Property Get iBRAND_BuyQuantity() As Long
    iBRAND_BuyQuantity = BuyQuantity
End Property
Private Property Get iBRAND_PreviousPrice() As Long
    iBRAND_PreviousPrice = PreviousPrice
End Property
Private Property Get iBRAND_PreviousRatio() As Double
    iBRAND_PreviousRatio = PreviousRatio
End Property
Public Property Get iBRAND_SellBuyRatio() As Double '買気配数量比率が大きいほど買い優勢とみる : 買気配値数量 / 合計気配数量
    If SellQuantity + BuyQuantity > 0 Then
        '買気配数量比率が大きいほど買い優勢とみる : 買気配値数量 / 合計気配数量
        iBRAND_SellBuyRatio = Format(BuyQuantity / (SellQuantity + BuyQuantity), "0.00")
    End If
End Property
Public Property Get iBRAND_TheoricPrice() As Double
    If iBRAND_SellBuyRatio = 0 Then Exit Property

    Dim PriceDiff As Long '売気配値と買気配値の価格差
    PriceDiff = SellPrice - BuyPrice
    
    If PriceDiff = 0 Then
        '板寄せ中
        TheoricPrice = SellPrice 'Me.AskingPriceでもよい（同一価格）
    Else
        TheoricPrice = Format(iBRAND_SellBuyRatio * PriceDiff + BuyPrice, "0.00")
    End If

End Property
'===================================


Private Sub mySheet_Calculate()
    If Val(mySheet.Range("B2").Value) = 0 Then Exit Sub
    
    BrandName = mySheet.Range("B1").Value
    Price = Val(mySheet.Range("B2").Value)
    TimeStamp = mySheet.Range("B3").Value
    SellPrice = Val(mySheet.Range("B4").Value)
    SellQuantity = Val(mySheet.Range("B5").Value)
    BuyPrice = Val(mySheet.Range("B6").Value)
    BuyQuantity = Val(mySheet.Range("B7").Value)
    PreviousPrice = Val(mySheet.Range("B8").Value)
    PreviousRatio = Val(mySheet.Range("B9").Value)
    
    RaiseEvent Change(Me)
End Sub
