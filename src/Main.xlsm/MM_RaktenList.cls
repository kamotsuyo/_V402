VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_RaktenList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents IDList As Worksheet
Attribute IDList.VB_VarHelpID = -1
Public WithEvents OList As Worksheet
Attribute OList.VB_VarHelpID = -1
Public WithEvents OListAlive As Worksheet
Attribute OListAlive.VB_VarHelpID = -1
Public WithEvents PList As Worksheet
Attribute PList.VB_VarHelpID = -1
Public WithEvents EXList As Worksheet
Attribute EXList.VB_VarHelpID = -1

'ファイルログ用
'Private myOlistLog As PrintLog

'CloseOrder自動発行用
Private myDicBackOrder As Dictionary  'key:OrderNO, Item : ORDER
Private myCollPosition As Collection  'Plistの更新にはユニークキーが存在しないのでコレクションとする

'ユーザーイベントを定義
Event IDlistChange(SerialNO_ As String, OrderNO_ As String, Result_ As String)
Event OlistChange(OrderNO_ As String, Price_ As String, Status_ As String, StopStatus_ As String, StopCondition_ As String)

Private Sub Class_Initialize()
    Call Init
End Sub

Private Sub Init()

    '---------------
    'IDList
    '---------------
    Set IDList = getMySheet("IDList")
    '項目を指定
    With IDList
        .Cells(2, 1) = "発注ID"
        .Cells(2, 2) = "注文番号"
        .Cells(2, 3) = "発注結果"
        .Cells(1, 1) = "=RssOrderIDList(R[1]C[0]:R[1]C[2] )"
'        .Visible = xlSheetHidden '非表示
    End With

    '---------------
    'OList
    '---------------
    Set OList = getMySheet("OList")
    With OList
        .Cells(2, 1) = "注文番号"
        .Cells(2, 2) = "取引"
        .Cells(2, 3) = "注文数量"
        .Cells(2, 4) = "注文単価"
        .Cells(2, 5) = "注文区分"
        .Cells(2, 6) = "通常注文状況"
        .Cells(2, 7) = "逆指値注文状況"
        .Cells(2, 8) = "逆指値条件"
        .Cells(2, 9) = "銘柄コード"
        .Cells(1, 1) = "=RssFOPOrderList(R[1]C[0]:R[1]C[8])"
'        .Visible = xlSheetHidden '非表示
    End With
    
   
    '---------------
    'OListAlive
    '---------------
    Set OListAlive = getMySheet("OListAlive")
    With OListAlive
        .Cells(2, 1) = "注文番号"
        .Cells(2, 2) = "取引"
        .Cells(2, 3) = "通常注文状況"
        .Cells(2, 4) = "注文数量"
        .Cells(2, 5) = "銘柄コード"
        .Cells(1, 1) = "=RssFOPOrderList(R[1]C[0]:R[1]C[4] , 1)"    '引数1指定：取消訂正可能
'        .Visible = xlSheetHidden '非表示
    End With
    
    
    '---------------
    'PList
    '---------------
    Set PList = getMySheet("PList")
    With PList
        .Cells(1, 1) = "=RssFOPPositionList()"
'        .Visible = xlSheetHidden '非表示
    End With
    '---------------
    'EXList
    '---------------
    Set EXList = getMySheet("EXList")
    With EXList
        .Range("a1").Formula = "=RssFOPExecutionList()"
'        .Visible = xlSheetHidden '非表示
    End With
    
    'ログ用インスタンス生成
'    Set myOlistLog = New PrintLog
'    Call myOlistLog.Init("Olist")   'ディレクトリ名称を指定
        
    
    'ディクショナリ・コレクションのインスタンス生成
    Set myDicBackOrder = New Dictionary
    Set myCollPosition = New Collection
End Sub



'--------------
'OrderNOとSerialNOを紐づけする
'結果にエラーがあればオーダーロックを解除する
Private Sub IDList_Change(ByVal Target As Range)
    If isDataRange(Target) = False Then Exit Sub

    Dim arr As Variant
    arr = Target
    
    Dim SerialNO As String
    Dim OrderNO As String
    Dim Result As String

    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        SerialNO = arr(i, 1)
        OrderNO = arr(i, 2)
        Result = arr(i, 3)
        
        '2度受信される
        '1回目はシリアルナンバーのみ受信：このときはOrderNOは"-",Resultは""である
        '２回目でシリアルナンバー、オーダーNO、リザルトがすべて受信される
        Debug.Print "MM_RakutenList IDList_Change"; SerialNO, OrderNO, Result
        If Result = "" Then
            'Resultが未取得の場合は１度目の受信なのでなにもしない
        Else
            '--------------
            'イベント発火
            '--------------
            RaiseEvent IDlistChange(SerialNO, OrderNO, Result)
        End If
        
        
        'リザルトエラー
        If Result <> "" And Result <> "発注済み" Then FormLog.addListItem (Result)
        
    Next i
End Sub


Private Sub OList_Change(ByVal Target As Range)
    If isDataRange(Target) = False Then Exit Sub  'データは増加のみ
    
    'ステータス保持：ステータス変更判定に使用
    Static PreviousStatus As String
    
    Dim OrderNO As String
    Dim BuySell As String
    Dim Quantity As String
    Dim Price As String
    Dim OrderType As String
    Dim Status As String
    Dim StopStatus As String
    Dim StopCondition As String
    Dim BrandCode As String
       
    Dim arr As Variant
    arr = Target
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        'ステータス変更時にイベント発行する
        OrderNO = arr(i, 1)
        BuySell = arr(i, 2)
        Quantity = arr(i, 3)
        Price = arr(i, 4)
        OrderType = arr(i, 5)
        Status = arr(i, 6)
        StopStatus = arr(i, 7)
        StopCondition = arr(i, 8)
        BrandCode = arr(i, 9)
        
        '-------------
        'イベント発火
        '-------------
        RaiseEvent OlistChange(OrderNO, Price, Status, StopStatus, StopCondition)
        
       
        '非同期でログ出力
'        Call myOlistLog.printAsync(DataArray(Arr, I))   '引数：該当の行[i]を１次元配列にしたもの
    Next i
End Sub


'ポジション
'クローズオーダー発行する
'基準マージン・基準ロスを使用して価格条件を決める(MainModuleで定義している)
Private Sub PList_Change(ByVal Target As Range)

    '-----増減データ判定方法-----
    '1)エンドマーク受信時にはStatic変数に格納して保持しておく
    '2)受信データの行を保持したエンドマーク行と比較して新規データかどうかを判定する
    If Target.Row < 3 Then Exit Sub '(Target.Rowは受信データレンジの開始行を表す)
    Static EndmarkRow As Integer
    If EndmarkRow = 0 Then EndmarkRow = 3 '初期値
    Dim CurrentRow As Integer '受信中のデータ行
    '-----------------------------
    
    Dim BrandCode As String
    Dim BuySell As String
    Dim OpenDate As String
    Dim OpenQuantity  As String
    Dim ClosedQuantity As String
    Dim OpenPrice  As String

    Dim arr As Variant
    arr = Target
    
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)

        CurrentRow = Target.Row + (i - 1) 'これが現在の受信中データ行である(データ行数を付加する[i]はArrが１オリジンなので-1している)

        Select Case arr(i, 1)    '１列目のセルで判定
            Case ""
                '削除行の受信
                
            Case ENDMARK
                'エンドマーク行の受信
                EndmarkRow = CurrentRow
            Case Else
                '新規ポジションか既存ポジションの更新か判定する
                '受信データ行がエンドマーク行以上なら新規データ受信
                If CurrentRow >= EndmarkRow Then
                    '----------------------------
                    '新規ポジションが追加された
                    '----------------------------
                    BrandCode = arr(i, 1)
                    BuySell = arr(i, 4)
                    OpenDate = arr(i, 5)
                    OpenQuantity = arr(i, 6)
                    ClosedQuantity = arr(i, 7)
                    OpenPrice = arr(i, 8)
                    
                    '-----------------------
                    'クローズオーダー自動発行
                    '-----------------------
                    Call CheckBackOrder(BrandCode, OpenDate, OpenPrice)
                    
                Else
                    '----------------------------
                    '既存ポジションが更新された
                    '----------------------------
                End If
                
        End Select
    Next i

End Sub


Private Sub EXList_Change(ByVal Target As Range)
    If isDataRange(Target) = False Then Exit Sub  'データは増加のみ

    

End Sub

'----------------
'CloseOrderを自動生成する
'----------------
'Olist主体でCloseOrderを生成する
'直接のMarketSpeedの注文には反応しない
'リストイベントの発生順を以下の順序と判断している
'1)Olistの更新：約定のステータス受信
'2)Plistの新規データ発生
'Plistのみの新規データ発生には反応しない：直接のMarketSpeedの注文には反応しない

'---------------
'setBackOrder
'Olistの更新時にStatusが約定ならばDEALクラスからCallされる
Public Sub setBackOrder(Order_ As iORDER)
    If Not myDicBackOrder.Exists(Order_.OrderNO) Then
        Call myDicBackOrder.Add(Order_.OrderNO, Order_)
        Debug.Print "addBackOrder:DicBackOrderを追加しました OrderNO="; Order_.OrderNO
    End If
End Sub

Private Sub CheckBackOrder(BrandCode_ As String, OpenDate_ As String, OpenPrice_ As String)
    If myDicBackOrder.Count = 0 Then Exit Sub
    
    Dim myOrderNO As String 'DicBackOrderの格納キーでもある
    Dim myOrder As SubProject.iORDER

    Dim i As Integer
    For i = 0 To myDicBackOrder.Count - 1 'ディクショナリは０オリジン :古い順に走査
        'BrandCodeが一致していればクローズオーダーを発行する
        myOrderNO = myDicBackOrder.Keys(i)
        Set myOrder = myDicBackOrder.Item(myOrderNO)
        
        'BrandCodeが一致しているか
        If myOrder.myStra.Brand.BrandCode = BrandCode_ Then
            
            '-----------------------
            'クローズオーダーを発行する
            '-----------------------
            Debug.Print "クローズオーダーを発行する", myOrderNO
            'V302版
'            Call myOrder.PublishCloseWithStop(OpenDate_, OpenPrice_)

            'V303版
            'V303ではオープンオーダー生成とクローズオーダー生成を分離
            Call OrderModule.CreateNewPosition(myOrder, OpenDate_, OpenPrice_)
'            Call createNewCloseOrder(myOrder, OpenDate_, OpenPrice_)
            
            '-----------------
            'myDicBackOrderディクショナリから削除
            '-----------------
            Call myDicBackOrder.Remove(myOrderNO)
            'Forループから抜ける
            Exit For
        Else
            'BrandCodeが一致していない
            'なにもしない
        End If
    Next i
End Sub

'Private Sub createNewCloseOrder(Order_ As ORDER, OpenDate_ As String, OpenPrice_ As String)
'    '------------------------
'    '新規クローズオーダー生成
'    '------------------------
'    Dim myCloseOrder As CloseOrder
'    Set myCloseOrder = New CloseOrder
'    Call myCloseOrder.Init(Order_, OpenDate_, OpenPrice_)
'
'    '----------------------
'    '逆指値付通常注文を発行
'    '----------------------
'    Call myCloseOrder.PublishCloseWithStop 'クローズオーダー発行に必要なプロパティ(OpenDate,OpenPrice)は上記でセット済のオープンオーダーインスタンスのものを使う
'
'End Sub




'全列データを配列に
'渡された２次元配列の指定行（index_）の列データを配列にして返す
'**元がRangeなので1オリジンになる
Private Property Get DataArray(Arr_ As Variant, Index_ As Integer) As Variant

    Dim myArr() As Variant
    ReDim myArr(LBound(Arr_, 2) To UBound(Arr_, 2))
    Dim j As Integer
    For j = LBound(Arr_, 2) To UBound(Arr_, 2)
        myArr(j) = Arr_(Index_, j)
    Next j
    DataArray = myArr
    
End Property

