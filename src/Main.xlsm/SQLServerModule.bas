Attribute VB_Name = "SQLServerModule"
Option Explicit

'事前に参照設定（Microsoft ActiveX Data Object 6.1 Library）

Const CONNSTRING As String = "Driver={MySQL ODBC 8.0 Unicode Driver};" & _
                          "Server=localhost;" & _
                          "Database=rakuten_rss;" & _
                          "User=kamouser;" & _
                          "Password=kam1824;" & _
                          "Charset=sjis;"

Public Function ConnectToMysqlDatabase_Select(sql As String) As Recordset
    Debug.Print sql
    'ADOを使用してMySQLに接続
    Dim cn As New ADODB.Connection
    cn.ConnectionString = CONNSTRING
    cn.Open

    'SELECT文の実行（取得した内容の確認）
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient 'ADO Recordset で RecordCount プロパティが -1 を返す場合の対策
    rs.Open sql, cn 'SQL文の実行


    'メモリの解放（無くとも構わない）
'    rs.Close: Set rs = Nothing
'    cn.Close: Set cn = Nothing
    
    Set ConnectToMysqlDatabase_Select = rs
    
End Function

Sub ConnectToMysqlDatabase_InsertInto(sql As String)
    'ADOを使用してMySQLに接続
    Dim cn As New ADODB.Connection 'Connectionオブジェクトのインスタンスを生成
    cn.ConnectionString = CONNSTRING
    
    cn.Open

    'INSERT文の実行
    Dim cm As New ADODB.Command 'Commandオブジェクトのインスタンスを生成
    cm.ActiveConnection = cn
    'cm.CommandText = "INSERT INTO test (name,age) VALUES ('awe',45) , ('kakak',55) "
    cm.CommandText = sql
    
    Dim Result As Long
    
    cm.Execute Result   '実行結果をresultに格納
    
    Debug.Print Result & "件追加しました"

    'メモリの解放
    Set cm = Nothing
    cn.Close: Set cn = Nothing
End Sub

Sub ConnectToMysqlDatabase_UPDATE(sql As String)
    'ADOを使用してMySQLに接続
    
    Debug.Print sql
    
    Dim cn As New ADODB.Connection 'Connectionオブジェクトのインスタンスを生成
    cn.ConnectionString = CONNSTRING
    cn.Open

    'UPDATE文の実行
    Dim cm As New ADODB.Command 'Commandオブジェクトのインスタンスを生成
    cm.ActiveConnection = cn
'    cm.CommandText = "UPDATE test SET age = 20 WHERE id = 3"
    cm.CommandText = sql
    
    Dim Result As Long
    cm.Execute Result
    
    Debug.Print Result & "件更新しました"
    
    'メモリの解放
    Set cm = Nothing
    cn.Close: Set cn = Nothing
End Sub

Public Function ConnectToMysqlDatabase_Duplicate(sql As String) As String
    'ON DUPLICATE KEY UPDATE を使用した場合、行ごとの影響を受けた行の戻り値は、
    '既存の行がその現在の値に設定された場合 : 0
    'その行が新しい行として挿入された場合 : 1
    '既存の行が更新された場合 : 2
    
    'test
'    Debug.Print sql
    
    Static arr() As Variant
    arr = Array("変更なし", "挿入", "更新")     '結果戻り値に対応する文字配列
    
    'ADOを使用してMySQLに接続
    Dim cn As New ADODB.Connection 'Connectionオブジェクトのインスタンスを生成
    cn.ConnectionString = CONNSTRING
    cn.Open

    'UPDATE文の実行
    Dim Result As Long
    Dim cm As New ADODB.Command 'Commandオブジェクトのインスタンスを生成
    cm.ActiveConnection = cn
    cm.CommandText = sql
    
    cm.Execute Result
    
    ConnectToMysqlDatabase_Duplicate = arr(Result)      '結果を配列の該当文字列として返す
    
    'メモリの解放
    Set cm = Nothing
    cn.Close: Set cn = Nothing
End Function

'新規orderIDの発番
'2023/5/10　修正：strategyNoを格納するように変更する
'2023/5/19   修正：もとに戻す。すべてnull値
Public Function GetOrderIdFromDB_T_orderIDList() As Long
'Public Function GetOrderIdFromDB_T_orderIDList(strategyNO As Integer) As Long
    'ADOを使用してMySQLに接続
    Dim cn As New ADODB.Connection 'Connectionオブジェクトのインスタンスを生成
    cn.ConnectionString = CONNSTRING
    
    cn.Open

    'INSERT文の実行
    Dim cm As New ADODB.Command 'Commandオブジェクトのインスタンスを生成
    cm.ActiveConnection = cn
    cm.CommandText = "INSERT INTO t_orderIDList () VALUES () ;"      'すべてnullでinsert
'    cm.CommandText = "INSERT INTO t_orderIDList (strategyno) VALUES ('" & strategyNO & "') ;"       'ストラテジ番号をinsert　　2023/5/10変更
    
    Dim Result As Long
    cm.Execute Result   '実行結果をresultに格納
    Debug.Print "orderID予約の空白行を" & Result & "件追加しました"
    
    '------------------------------
    'SELECT文の実行（orderIDの確認）
    Dim sql As String
    sql = "select max(orderid) from t_orderidlist ;"
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient 'ADO Recordset で RecordCount プロパティが -1 を返す場合の対策
    rs.Open sql, cn 'SQL文の実行
    
    Debug.Print rs(0)
    
    GetOrderIdFromDB_T_orderIDList = rs(0)

    'メモリの解放
    Set rs = Nothing
    Set cm = Nothing
    cn.Close: Set cn = Nothing
End Function

