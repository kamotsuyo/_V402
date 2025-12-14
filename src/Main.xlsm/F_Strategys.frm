VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Strategys 
   Caption         =   "UserForm1"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13725
   OleObjectBlob   =   "F_Strategys.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F_Strategys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
'**********************
'CollStraを走査して登録済みのストラテジ一覧を作成する
'**********************

Private myDicListItem As Dictionary 'Key:myStra.name , Item:ListItem

Private Sub UserForm_Initialize()
    Me.Caption = TypeName(Me)
    
    'リストビュー　ヘッダーを生成
    With Me.ListView1.ColumnHeaders
        .Add 1, , "Name"
        .Add 2, , "Market"
        .Add 3, , "AutoExit"
        .Add 4, , "更新時刻"
        .Add 5, , "OrderLock"
        .Add 6, , "AliveCount"
        .Add 7, , "TotalCount"

        .Add 8, , "TotalPL"
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

Public Sub Init()
    'ListItemを宣言
    Dim myListItem As ListItem

    Dim myStra As iSTRA
    For Each myStra In CollStra
        
        If myDicListItem.Exists(myStra.Name) Then
            '既存
            Set myListItem = myDicListItem.Item(myStra.Name)
        Else
            '新規
            Set myListItem = Me.ListView1.ListItems.Add 'Add引数なし：一番下に追加
            'ディクショナリに格納する
            Call myDicListItem.Add(myStra.Name, myListItem)
        End If
        
        'Text
        myListItem.Text = myStra.Name
        myListItem.SubItems(1) = myStra.Brand.PetName
        myListItem.SubItems(2) = TypeName(myStra.AutoExit)

    Next
End Sub


Public Sub Display(Stra_ As iSTRA)
    'ListItemを宣言
    Dim myListItem As ListItem
    Set myListItem = myDicListItem.Item(Stra_.Name)
    
    '-----------
    'SubItem列を記述
    With myListItem

        .SubItems(3) = Format(Now, "hh:nn:ss")
        .SubItems(4) = Stra_.OrderLock
        .SubItems(5) = Stra_.AliveCount
        .SubItems(6) = Stra_.TotalCount
        .SubItems(7) = Stra_.TotalProfitLoss
    End With
    
End Sub
