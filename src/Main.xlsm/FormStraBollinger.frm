VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormStraBollinger 
   Caption         =   "UserForm1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14700
   OleObjectBlob   =   "FormStraBollinger.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormStraBollinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit

Private DicListItem As Dictionary 'Key:SerialNO , Item:Array

Private myStra As iSTRA  '生成元のストラテジ
Private myDummyLog As PrintLogDummyOrder

Private TotalCommission As Long '手数料合計
Private TotalProfitLoss As Long '損益合計

Const COMMISSHION As Integer = -77
Const PIPSQUANTITY As Integer = 100 '225miniの1pipsの数量

Private Sub UserForm_Initialize()
    
    'ヘッダーを指定する
    Dim ArrHeader As Variant
    ArrHeader = Array("SerialNO", "時刻", "売買", "価格", "エントリー時パラメータ", "更新時データ", "Exit", "損益")
    Call SubModule.setListViewHeader(Me.ListView1, ArrHeader)

    With Me.ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .Font.Size = 10
        .Font = "Meiryo UI"
        .ColumnHeaders.Item(1).Width = 60 'ヘッダー生成後に
        .ColumnHeaders.Item(2).Width = 60 'ヘッダー生成後に
        .ColumnHeaders.Item(3).Width = 30
        .ColumnHeaders.Item(4).Width = 60
        .ColumnHeaders.Item(5).Width = 200
        .ColumnHeaders.Item(6).Width = 200
        .ColumnHeaders.Item(7).Width = 60
        .ColumnHeaders.Item(8).Width = 60
    End With
    
    'ディクショナリをインスタンス生成
    Set DicListItem = New Dictionary
    
End Sub

'生成元のストラテジ
Public Sub Init(Stra_ As iSTRA)
    Set myStra = Stra_
    Me.Caption = TypeName(myStra)   'キャプション指定
    
    'ファイルログ
    Set myDummyLog = New PrintLogDummyOrder
    Call myDummyLog.Init(TypeName(myStra) & "_Exit")
End Sub

Public Sub insertDummyOrder(SerialNO_ As String, DummyOrder_ As DummyOrder)

    'ListItemを定義しておく：新規、既存ともに使用する
    Dim myListItem As ListItem
    If Not DicListItem.Exists(SerialNO_) Then
        '新規
        'リストアイテムの生成
        Set myListItem = Me.ListView1.ListItems.Add(1) '引数１を指定して一番上に追加する
        myListItem.Text = SerialNO_ 'リストアイテムのテキストをセットする
        Call DicListItem.Add(SerialNO_, myListItem) 'DicListItemディクショナリに格納する
        
        '新規オーダーなら
        '手数料増加
        TotalCommission = TotalCommission + COMMISSHION
    Else
        '既存
        Set myListItem = DicListItem.Item(SerialNO_) 'ディクショナリから既存のListItemを取得
        'リストアイテムのテキストは変わらないので指定不要
    End If
    
    '共通処理
    'リストアイテムのサブアイテムへセットする
    With myListItem
        .SubItems(1) = DummyOrder_.myTime
        .SubItems(2) = DummyOrder_.BuySell
        .SubItems(3) = DummyOrder_.OpenPrice
        .SubItems(4) = Join(DummyOrder_.OpenArgs, "_")
    End With
    
    'このフォームの表示を更新する
    Call updateForm
End Sub

Public Sub updateDummyOrder(SerialNO_ As String, UpdateArgs_ As Variant)
   
    Dim myListItem As ListItem
    Set myListItem = DicListItem.Item(SerialNO_)
    
    myListItem.SubItems(5) = Join(UpdateArgs_, "_")
End Sub

Public Sub ExitDummyOrder(SerialNO_ As String, ExitArgs_ As Variant)
    
    Debug.Print "exit", TypeName(myStra), SerialNO_, ExitArgs_(0), Val(ExitArgs_(1))
    
    Dim myListItem As ListItem
    Set myListItem = DicListItem.Item(SerialNO_)
    
    myListItem.SubItems(6) = ExitArgs_(0)   'Exit時刻
    myListItem.SubItems(7) = ExitArgs_(1)   'Exit利益(pips)
    
    '損益計算
    TotalProfitLoss = TotalProfitLoss + Val(ExitArgs_(1))
    
    'このフォームの表示を更新する
    Call updateForm
    
    
    '----------
    'ログ出力
    '----------
    Dim ArrLog As Variant
    With myListItem
        ArrLog = Array( _
                    "[" & .Text & "]", _
                    .SubItems(1), _
                    .SubItems(2), _
                    .SubItems(3), _
                    .SubItems(4), _
                    .SubItems(5), _
                    .SubItems(6), _
                    .SubItems(7), _
                    Me.TextBox3, _
                    Me.TextBox4, _
                    Me.TextBox5 _
                    )
    End With
    

    'ファイルログ
    Call myDummyLog.printAsync(ArrLog)
    
End Sub

Private Property Get ArrAllSubitems(ListItem_ As ListItem) As Variant
    Dim arr As Variant
    ReDim arr(1 To 7)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
    
        arr(i) = CStr(ListItem_.SubItems(i))

    Next i
    ArrAllSubitems = arr
End Property

Private Sub updateForm()
    With Me
        .TextBox3 = TotalProfitLoss 'トータル損益
        .TextBox4 = TotalCommission '合計手数料]
        .TextBox5 = TotalProfitLoss * PIPSQUANTITY + TotalCommission  '損益合計-手数料合計
    End With
End Sub
