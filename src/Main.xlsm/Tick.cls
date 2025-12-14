VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'4/2動作確認済
'#####################
Implements iTECBASE
'#####################

'発行するイベント
Event Change(Tick_ As iTECBASE)
'受信イベント
Private WithEvents mySheet As Worksheet
Attribute mySheet.VB_VarHelpID = -1

'ディクショナリ
Private DicCandle As Dictionary
'Candle用定数
Const ASHI As Integer = 100    '取得本数：MarketSpeed仕様最大300だが速度を考慮して100にしておく
'Const MAXCOUNT As Integer = 30 'DicTickCandleに格納する最大件数

Private Number As Integer 'まとめる単位　**この設定は{10,20,30など}60以下でなくてはならない

'処理時間計測
Private El As ElapsedTime

Private Sub Class_Initialize()
    'ディクショナリインスタンス生成
    Set DicCandle = New Dictionary
    '処理時間計測
    Set El = New ElapsedTime
End Sub

'Ver0606-2:まとめ単位をConstでなく引数からセットするよう変更
'ver0606-3:このままでmySheetを共有して複数のまとめ単位を異なるTickを生成できる
Public Sub Init(BrandCode_ As String, Optional Number_ As Integer = 15)
    'Tickの受信用ワークシートを生成（生成済みならそれを取得）
    Set mySheet = getMySheet("T_" & BrandCode_)
    mySheet.Range("a1").Formula = "=RssTickList( , " & BrandCode_ & "," & ASHI & ")"
    mySheet.Visible = xlSheetHidden '非表示
    
    'まとめ単位をセットする
    Number = Number_
End Sub


'==========================
'TECHINICALBASE共通プロパティ
Private Property Get iTECBASE_Candle(Optional I_ As Integer = 0) As iTECBASECANDLE
    If DicCandle.Count > I_ Then
        Set iTECBASE_Candle = DicCandle.Items(DicCandle.Count - 1 - I_)
    End If
End Property
Private Property Get iTECBASE_MostHighCandle(Optional N_ As Integer = 5, Optional I_ As Integer = 0) As iTECBASECANDLE
    Dim myCandle As iTECBASECANDLE
    
    'Candleプロパティを使う
    Dim i As Integer
    For i = I_ To I_ + N_ - 1   '開始位置を示す引数I_を使う
        If i > DicCandle.Count - 1 Then Exit For    '格納済みキャンドルが不足していれば抜ける
        Set myCandle = Candle(i)
        If myCandle.Volume > 0 Then '出来高０の場合は処理対象外
            If iTECBASE_MostHighCandle Is Nothing Then
                Set iTECBASE_MostHighCandle = myCandle
            ElseIf myCandle.HighPrice > iTECBASE_MostHighCandle.HighPrice Then  '同値なら新しい日時のデータのままにしておく
                Set iTECBASE_MostHighCandle = myCandle
            End If
        End If
    Next i
End Property
Private Property Get iTECBASE_MostLowCandle(Optional N_ As Integer = 5, Optional I_ As Integer = 0) As iTECBASECANDLE
    Dim myCandle As iTECBASECANDLE
    'Candleプロパティを使う
    Dim i As Integer
    For i = I_ To I_ + N_ - 1   '開始位置を示す引数I_を使う
        If i > DicCandle.Count - 1 Then Exit For    '格納済みキャンドルが不足していれば抜ける
        Set myCandle = Candle(i)
        If myCandle.Volume > 0 Then '出来高０の場合は処理対象外
            If iTECBASE_MostLowCandle Is Nothing Then
                Set iTECBASE_MostLowCandle = myCandle
            ElseIf myCandle.LowPrice < iTECBASE_MostLowCandle.LowPrice Then  '同値なら新しい日時のデータのままにしておく
                Set iTECBASE_MostLowCandle = myCandle
            End If
        End If
    Next i
End Property
Private Property Get iTECBASE_Count() As Integer
    iTECBASE_Count = DicCandle.Count
End Property
Private Property Get iTECBASE_PriceRCI(Optional N_ As Integer = 10) As Double
    iTECBASE_PriceRCI = getRCI(getArrPrice(Me, N_))
End Property
Private Property Get iTECBASE_PriceRSI(Optional N_ As Integer = 10) As Double
    'RSIの期間１０であれば、差を１０件分計算するので、サンプル件数は１１件必要である。
    iTECBASE_PriceRSI = getRSI(getArrPrice(Me, N_ + 1))
End Property
Private Property Get iTECBASE_Ema5() As Double
    iTECBASE_Ema5 = getEMA5(Me)
End Property
Private Property Get iTECBASE_Ema25() As Double
    iTECBASE_Ema25 = getEMA25(Me)
End Property
Private Property Get iTECBASE_Ema5RSI(Optional N_ As Integer = 5) As Double
    iTECBASE_Ema5RSI = getRSI(getArrEMA5(Me, N_))
End Property
Private Property Get iTECBASE_VolumeRSI(Optional N_ As Integer = 5) As Double
    iTECBASE_VolumeRSI = getRSI(getArrVolume(Me, N_))
End Property
'==========================

'日時を比較するためにDate型に変換する
'Date型の注意点：時刻が0埋めにならない。ディクショナリに格納するときにはString型に変換して
'CStr(Format(myDate, "yyyy/mm/dd hh:nn:ss"))で格納すること。
Private Function getCurrentDateTime(Time_ As String) As Date
    '引数のTime_ はHH:NN:SS形式である
'    Debug.Print Hour(Time_), Date & " " & Time_
    '6/13修正版
    
    '処理時間計測：0.017ミリ秒程度
'    Call El.SetLap
        
    Dim myDateTime As Date
    myDateTime = Date & " " & Time_ 'これでDate型になる

    If myDateTime - Now > 0 Then
        myDateTime = DateAdd("d", -1, myDateTime)
    End If
    
    getCurrentDateTime = myDateTime
    
End Function

Private Function getDateTimeClass(DateTime_ As Date) As Date
    '時刻秒に変換する
    '処理時間計測：0.017ミリ秒程度
'    Call El.SetLap
    Dim Sec As Integer
    Sec = Second(DateTime_)
    Sec = Int(Sec / Number) * Number 'まるめ処理

    getDateTimeClass = Format(DateTime_, "yyyy/mm/dd hh:nn") & ":" & Format(Sec, "00")
    '下記は0.011ミリ秒程度
'    getDateTimeClass = Hour(DateTime_) & ":" & Minute(DateTime_) & ":" & Format(Sec, "00")
'    Debug.Print El.GetLap
End Function

'格納済みのCandleを返す
'引数Index_を指定しないときは最新足を返す
'N本前の足を取得する場合は引数に「N」を指定すればよい (1本前なら1,24本前なら24)
Public Property Get Candle(Optional Index_ As Integer = 0) As TickCandle
    If DicCandle.Count > Index_ Then
        Set Candle = DicCandle.Items(DicCandle.Count - 1 - Index_)
    End If
End Property

Private Property Get NewIndex() As Integer
    Static Counter As Integer
    Counter = Counter + 1
    NewIndex = Counter
End Property


Private Sub mySheet_Change(ByVal Target As Range)
    If isDataRange(Target) = False Then Exit Sub
    
    
    '最新データの日時
'    CurrentDateTime = getCurrentDateTime(mySheet.Cells(3, 1))   '3行目の1列目が一番新しいデータの時刻
    
    Dim LastCandle As TickCandle
    Set LastCandle = Candle 'Changeイベント受信時点での最後のTickCandleを格納する：Forループによって更新される前に。
    

    '処理時間計測
'    Call El.SetLap
    '更新範囲のデータ配列を作る
    '初期本数100の場合 1.5ミリ秒程度
    '初期本数300の場合 4ミリ秒程度
    Dim ArrData As Variant
    ArrData = getArrUpdateData(Target) '配列として格納されている：{日時,Volume,Price}
    If IsEmpty(ArrData) Then Exit Sub
    
    '処理時間計測
'    Debug.Print "更新範囲のデータ配列生成", El.GetLap
    
    

    '初期本数100の場合 4.5ミリ秒程度
    '初期本数300の場合 17ミリ秒程度
    'まとめ秒ごとのTickCandleを生成してDicCandleに格納していく。(格納済みの場合は更新する)
    Dim myDateTimeClass As Date
    Dim myCandle As TickCandle
    Dim i As Integer
    For i = UBound(ArrData) To LBound(ArrData) Step -1 '逆順で走査：古い時刻順に
        
        myDateTimeClass = getDateTimeClass(CDate(ArrData(i)(0)))

        If DicCandle.Exists(myDateTimeClass) Then
            '既存
            Set myCandle = DicCandle.Item(myDateTimeClass)
        Else
            '新規
            Set myCandle = New TickCandle
            Call myCandle.Init(myDateTimeClass, NewIndex)
            
            'ディクショナリに格納
            Call DicCandle.Add(myDateTimeClass, myCandle)
            '規定格納数を超えたら一番過去の足を削除
'            If DicCandle.Count > MAXCOUNT Then Call DicCandle.Remove(DicCandle.Keys(0))
        End If
        
        '共通処理
        Call myCandle.addData(CDate(ArrData(i)(0)), Val(ArrData(i)(1)), Val(ArrData(i)(2)))

        '処理時間計測
'        Call El.SetLap
        '処理時間 0.05ミリ秒程度
        '6/20 SMAの計算を高速化：
        'ミリ秒(0.022)
        'ミリ秒(0.017)
        
        With myCandle
            .SMA5 = getSMA5(Me)
            .SMA25 = getSMA25(Me)
            .EMA5 = getEMA5(Me)
            .EMA25 = getEMA25(Me)
        End With
        '処理時間計測
'        Debug.Print El.GetLap
        
    Next i
    


    '---------------
    'イベント発火
    '---------------
    RaiseEvent Change(Me)
End Sub

Private Function getArrUpdateData(Target_ As Range) As Variant
    On Error GoTo ErrorHandler
    
    Static PreviousDT As Date
    
    Dim LastDateTime As Date    '前回の最終日時を格納：初期値は"00:00:00"
    If Not iTECBASE_Candle Is Nothing Then LastDateTime = iTECBASE_Candle.LastDateTime
    
'    Debug.Print "LastDateTime"; LastDateTime, "PreviousDT"; PreviousDT
    
    Dim arr As Variant
    arr = Target_

    Dim ArrResult As Variant
    ReDim ArrResult(LBound(arr) To UBound(arr))

    Dim myIndex As Integer '配列に格納する際のインデックス番号として利用する
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        '現在日時をセットする
        Dim myCurrentDateTime As Date
        myCurrentDateTime = getCurrentDateTime(CStr(arr(i, 1)))
        
        If myCurrentDateTime - LastDateTime >= 0 Then
            myIndex = myIndex + 1
            '新しい時刻か同一の時刻の場合：更新対象のデータなので配列に格納する
            '配列として格納する：日時,Volume,Price
            ArrResult(myIndex) = Array(myCurrentDateTime, arr(i, 2), arr(i, 3))
        Else
            
            '更新データのRedimで格納済みだけの配列を作る
            ReDim Preserve ArrResult(1 To myIndex - Candle.AddedCount)
            
            'Exitがなかった! 6/15追記
            Exit For

        End If
    Next i
    
    '前回日時を更新する
    PreviousDT = getCurrentDateTime(mySheet.Cells(3, 1))
    
    getArrUpdateData = ArrResult
Exit Function
ErrorHandler:
    Debug.Print ("Tick.getArrUpdateDataエラー : " & Err.Description)
    Debug.Print ("myIndex : " & myIndex & ",  Candle.AddedCount : " & Candle.AddedCount)
End Function

