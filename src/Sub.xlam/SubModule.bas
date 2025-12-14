Attribute VB_Name = "SubModule"
Option Explicit

Public Const ENDMARK As String = "--------"

'指定の引数をシート名としてワークシートを追加作成する
'シート名から既存かどうかを確認し、既存ならそのシートをセット。ないなら新規作成。
'**リスト系クラスで使用中

'2025/12/6 6:15初回起動時にエラーで停止。
'引数SheetNameにてエラー発生：以下の文字列がSheetNameに格納される
'マーケットスピード II にログインして「接続中」の状態に変更してから利用してください。
'原因：MarketSpeedが「未接続」の状態で「発注可」に切り替える動作をするとこのエラーが発生。
'対策====> MM_ConnectionクラスのMarketSpeedContolModuleの待機時間を延長する

Public Function getMySheet(SheetName As String, Wb_ As Workbook) As Worksheet
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Wb_.Sheets
        
        If CurrentSheet.Name = SheetName Then
            Set getMySheet = CurrentSheet
            Exit Function
        End If
    Next
    Set getMySheet = Wb_.Worksheets.Add
    getMySheet.Name = SheetName '<==エラー発生

End Function

'ワークシートのデータかどうかの判別を行い、データであればtrueを返す
Public Property Get isDataRange(ByVal Target As Range) As Boolean
    
    isDataRange = True '初期値
    
    '1)データ行の開始位置よりも下行かどうか
    Const DATASTARTROW As Integer = 3 'データの規定の開始位置
    If Target.Row < DATASTARTROW Then isDataRange = False
    
    '2)ターゲットがエンドマーク行であるかどうかを判定する
    'ターゲットの開始位置と終了位置のセル内容[Target.Item(1).Value]がエンドマークならばそのデータはエンドマーク行であると判定できる
    If Target.Item(1).Value = ENDMARK And Target.Item(Target.Count) = ENDMARK Then isDataRange = False
    
End Property

'=========================
'オーダー系のメソッド
'=========================
'シリアルNO発行
Public Property Get NewSerialString() As String
    '新規ディール発行時のシリアルNO：発行時に+1増分させる
    Static Serial As Integer
    Serial = Serial + 1

    NewSerialString = CStr(Serial) '戻り値
End Property

'========================
'シート・フォーム系
Public Sub DeleteSheetWithoutSheet1(Wb_ As Workbook)
    Application.DisplayAlerts = False 'アラート不要
  
    If Wb_.Worksheets.Count > 1 Then
        Dim Ws As Worksheet
        For Each Ws In Wb_.Sheets
            If Ws.Name = "Sheet1" Then
                'ワークシートクリア
                Ws.Cells.Clear
                GoTo Skip
            End If
            Ws.Delete
Skip:
        Next
    End If
End Sub

Public Sub DeleteSheetWithout(ArrSheetsNames As Variant)
    Application.DisplayAlerts = False 'アラート不要
    
    Dim Ws As Worksheet
    Dim SheetName As Variant
    For Each Ws In ThisWorkbook.Sheets
        For Each SheetName In ArrSheetsNames
            If Ws.Name = SheetName Then
                GoTo Skip
            End If
        Next
        Ws.Delete
Skip:
    Next
End Sub

Public Sub HideForms(Optional ExemptName_ As String)
    'Form_Mainを除いてHideする
    Application.DisplayAlerts = False   '注意を表示しない
    Dim i As Integer
    
    'UserFormsオブジェクトは0オリジン
    For i = UserForms.Count - 1 To 0 Step -1    '降順で走査
        If ExemptName_ = "" Then
            '全フォームを閉じる
            UserForms.Item(i).Hide
            
        Else
            '除外フォームあり
            If UserForms.Item(i).Name <> ExemptName_ Then UserForms.Item(i).Hide
        End If
    Next i
End Sub

Public Sub UnloadAllForms()
    '全フォームをアンロード
    Application.DisplayAlerts = False   '注意を表示しない
    Dim i As Integer
    
    For i = UserForms.Count - 1 To 0 Step -1
        'UserFormsオブジェクトは0オリジン
        '全フォームをアンロード
        Unload UserForms.Item(i)

    Next i
End Sub

Public Property Get isAliveStatus(Status_ As String) As Boolean
'ステータス一覧
'執行待ち   ==Alive
'執行中     ==Alive
'出来有
'約定
'取消中（出来有）
'取消中（出来無）
'取消済（出来有）
'取消済（出来無）
'出来ず（出来有）
'出来ず（出来無）
'訂正済     ==Alive

'の中で
'"執行待ち"
'"執行中"
'"訂正済"
'ならばAliveとする
    Select Case Status_
        Case "執行待ち", "執行中", "訂正済"
            isAliveStatus = True
        Case Else
            isAliveStatus = False
    End Select
End Property

'5円単位まるめ＆切り捨て 'FutureOrder,StockOrderで使用
Public Function GetRound5Down(Str_ As String) As String
    Const BASE As Integer = 5
    GetRound5Down = Int(Val(Str_) / BASE) * BASE
End Function
'5円単位まるめ＆切り上げ
Public Function GetRound5Up(Str_ As String) As String
    Const BASE As Integer = 5
    GetRound5Up = Int(Val(Str_) / BASE) * BASE + BASE
End Function
