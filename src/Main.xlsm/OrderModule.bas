Attribute VB_Name = "OrderModule"
Option Explicit
'*****************
'Main OrderModule
'V303
'*****************

Public Function getStraTotalCount(Stra_ As iSTRA) As Long
    Dim Counter As Long
    
    Dim TargetOrder As iORDER
    Dim Key As Variant
    For Each Key In DickOorder
        Set TargetOrder = DickOorder.Item(Key)
        '条件：ストラテジ名称が同一
        If TargetOrder.myStra.Name = Stra_.Name Then
            '現状オーダー件数を加算
            Counter = Counter + TargetOrder.Quantity
        End If
    Next
    getStraTotalCount = Counter
End Function

Public Function getStraAliveCount(Stra_ As iSTRA) As Long
    Dim Counter As Long
    
    Dim TargetOrder As iORDER
    Dim Key As Variant
    For Each Key In DickOorder
        Set TargetOrder = DickOorder.Item(Key)
        '条件：DEALがAliveである
        '条件：DEALのストラテジ名称が同一
        If TargetOrder.isAlive And TargetOrder.myStra.Name = Stra_.Name Then
            '現状オーダー件数を加算
            Counter = Counter + TargetOrder.Quantity
        End If
    Next
    getStraAliveCount = Counter
End Function


'ストラテジ内部で新規ディール発行の際に使用
Public Function getNewOrder(Stra_ As iSTRA, Quantity_ As Long) As iORDER
    '発行の可否判定
    '---------------
    '対象時間外かどうか
    If Stra_.isOutSideHours Then
        Set getNewOrder = Nothing
        'Log
        FormLog.addListItem Stra_.Name & ":対象時間外"
        Exit Function
    End If
    
    'オーダーロック中かどうか
    If Stra_.OrderLock Then
        Set getNewOrder = Nothing
        'Log
        FormLog.addListItem Stra_.Name & ":オーダーロック中"
        Exit Function
    End If
    
    'オーダーしようとしている数量と現状オーダー数を合わせると規定の最大発注可能を超えないか
    Dim CurrentOrderQuantity As Long
    '上記のメソッドgetStraAliveCountを使う
    CurrentOrderQuantity = getStraAliveCount(Stra_)
    
    '判定
    If CurrentOrderQuantity + Quantity_ > Stra_.MAXORDERQUANTITY Then
        Set getNewOrder = Nothing
        'Log
        FormLog.addListItem Stra_.Name & ":最大発注可能を超える"
        Exit Function
    End If
    
    '---------------------
    'ここまで来るとオーダー可である
    Debug.Print "オーダー可"
    'オーダー可の場合は
    'オーダーロック実行
    Call Stra_.changeOrderLock(True)
    
    
    '------------------
    '新規オーダーインスタンス生成
    '------------------
    Dim myOrder As iORDER
    'DummyOrder,FutureOrder,StockOrderの分岐
    'iSTRAのDealTypeプロパティを使って分岐する
    Select Case Stra_.DealType
        Case En_DealType.DummyOrder
            Set myOrder = New DummyOrder
        Case En_DealType.FutureOrder
            Set myOrder = New FutureOrder
        Case En_DealType.StockOrder
            Set myOrder = New StockOrder
    End Select
    
    Call myOrder.Init(Stra_, 1)
    'DicDealに追加格納 : DEALクラス内で行う
    
    '戻り値にセット
    Set getNewOrder = myOrder
    
End Function



'V303追加
'ストラテジ内部で新規ディール発行の際に使用
'MM_RaktenListのCheckBackOrderメソッドからCallされる
'戻り値なし
Public Sub CreateNewPosition(Order_ As iORDER, OpenDate_ As String, OpenPrice_ As String)
    '--------------
    '新規ポジションの生成(クローズオーダー)
    '--------------
    Dim myPosition As iPOSITION
    If TypeOf Order_ Is FutureOrder Then Set myPosition = New FuturePosition
    If TypeOf Order_ Is StockOrder Then Set myPosition = New StockPosition
    If TypeOf Order_ Is DummyOrder Then Set myPosition = New DummyPosition
    
    Call myPosition.Init(Order_, OpenDate_, OpenPrice_)

End Sub

