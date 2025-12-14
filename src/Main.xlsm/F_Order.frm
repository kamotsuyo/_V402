VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Order 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   OleObjectBlob   =   "F_Order.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private myDicListItem As Dictionary 'Key:SerialNO , Item:ListItem


Private Sub UserForm_Initialize()
    Me.Caption = TypeName(Me)
    
    'リストビュー　ヘッダーを生成
    With Me.ListView1.ColumnHeaders
        .Add 1, , "SerialNO"
        .Add 2, , "StraName"
        .Add 3, , "DealType"
        .Add 4, , "BuySell"
        .Add 5, , "Quantity"
        .Add 6, , "TimeStamp"
        .Add 7, , "isAlive"
        .Add 8, , "ProfitLoss"
        .Add 9, , "ExitTimeStamp"
    End With
    
    With Me.ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .Font.Size = 10
        .Font = "Meiryo UI"
        .ColumnHeaders.Item(1).Width = 50 'ヘッダー生成後に
        .ColumnHeaders.Item(2).Width = 50 'ヘッダー生成後に
    End With
    
    'ディクショナリインスタンス
    Set myDicListItem = New Dictionary
    
End Sub

Public Sub InsertDeal(Order_ As iORDER)
    'ListItemを宣言
    Dim myListItem As ListItem
    '新規
    Set myListItem = Me.ListView1.ListItems.Add 'Add引数なし：一番下に追加
    'ディクショナリに格納する
    Call myDicListItem.Add(Order_.SerialNO, myListItem)
    
    With myListItem
        .Text = Order_.SerialNO
        .SubItems(1) = Order_.myStra.Name
        'En_DealTypeを文字列にする
        'DummyOrder = 0,    FutureOrder = 1,    StockOrder = 2
        If TypeOf Order_ Is DummyOrder Then
            .SubItems(2) = "ダミー"
        ElseIf TypeOf Order_ Is FutureOrder Then
            .SubItems(2) = "先物"
        ElseIf TypeOf Order_ Is StockOrder Then
            .SubItems(2) = "株式"
        End If
'        Select Case Order_.DealType
'            Case 0
'                .SubItems(2) = "ダミー"
'            Case 1
'                .SubItems(2) = "株式"
'            Case 2
'                .SubItems(2) = "先物"
'        End Select
        
        'En_Buysellを文字列にする
        '売 = 1,    買 = 3
        Select Case Order_.BuySell
            Case 1
                .SubItems(3) = "売"
            Case 3
                .SubItems(3) = "買"
        End Select
        
        .SubItems(4) = Order_.Quantity
        .SubItems(5) = Order_.TimeStamp
        .SubItems(6) = Order_.isAlive
        .SubItems(7) = Order_.ProfitLoss
        .SubItems(8) = Order_.ExitTimeStamp
    End With
End Sub

Public Sub UpdateDeal(Order_ As iORDER)
    'ListItemを宣言
    Dim myListItem As ListItem
    '新規
    Set myListItem = myDicListItem.Item(Order_.SerialNO)
    
    With myListItem
        .SubItems(6) = Order_.isAlive
        .SubItems(7) = Order_.ProfitLoss
        .SubItems(8) = Order_.ExitTimeStamp
    End With
End Sub
