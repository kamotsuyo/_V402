VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSub 
   Caption         =   "UserForm1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5595
   OleObjectBlob   =   "FormSub.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = Me.Name
    
    'ヘッダーを生成
    Dim myListView As ListView
     'ヘッダーを生成
    Set myListView = Me.ListView1
    With myListView.ColumnHeaders
        .Add 1, , "時刻"
        .Add 2, , "ログ"
    End With
    
    With myListView
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .Font.Size = 10
        .Font = "Meiryo UI"
        .ColumnHeaders.Item(1).Width = 80 'ヘッダー生成後に
        .ColumnHeaders.Item(2).Width = 170 'ヘッダー生成後に
    End With
End Sub

Public Sub addListItem(Log_ As String)
'データを追加するときは、ListItemsコレクションのAddメソッドを実行します。
'注意しなければならないのは、ListViewコントロールを表形式(lvwReport)で使用する場合、
'ListItemsコレクションは左端列(1列目)の項目だということです。
'2列目以降は、1列目に登録したデータのSubItemsコレクションになります。

    'ListItemを生成する
    Dim myListItem As ListItem
    Set myListItem = Me.ListView1.ListItems.Add(1) '引数１を指定して一番上に追加する

    
    myListItem.Text = kernel32_v2.getTimeStamp
    myListItem.SubItems(1) = Log_
    
    myListItem.Selected = True  'フォーカス
End Sub
