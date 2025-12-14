VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLog 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   OleObjectBlob   =   "FormLog.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0  '初期位置を0:手動にセット
    
    'ヘッダーを生成
    With Me.ListView1.ColumnHeaders
        .Add 1, , "時刻"
        .Add 2, , "ログ"
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
        .ColumnHeaders.Item(1).Width = 150 'ヘッダー生成後に
        .ColumnHeaders.Item(2).Width = 600 'ヘッダー生成後に
    End With
    
End Sub

Public Sub addListItem(Log_ As String)
'データを追加するときは、ListItemsコレクションのAddメソッドを実行します。
'注意しなければならないのは、ListViewコントロールを表形式(lvwReport)で使用する場合、
'ListItemsコレクションは左端列(1列目)の項目だということです。
'2列目以降は、1列目に登録したデータのSubItemsコレクションになります。

    'ListItemを生成する
    Dim myListItem As ListItem
    Set myListItem = FormLog.ListView1.ListItems.Add(1) '引数１を指定して一番上に追加する

    
    myListItem.Text = kernel32_v2.getTimeStamp
    myListItem.SubItems(1) = Log_
    
    myListItem.Selected = True  'フォーカス
End Sub
