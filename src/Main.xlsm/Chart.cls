VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iTECBASE
'#####################

'発行するイベント
Event Change(Chart_ As iTECBASE)
Event PriceChange(Chart_ As iTECBASE)

'受信イベント
Private WithEvents mySheet As Worksheet
Attribute mySheet.VB_VarHelpID = -1

'ディクショナリ
Private DicCandle As Dictionary
'定数
Const ASHI As Integer = 500 '初期に読み込む足の本数：最大3000
'Const MAXCOUNT As Integer = 30 'DicChartに格納する最大件数

'処理時間計測
Private El As ElapsedTime

Private Sub Class_Initialize()
    'ディクショナリインスタンス生成
    Set DicCandle = New Dictionary
    '処理時間計測
    Set El = New ElapsedTime
End Sub

Public Sub Init(BrandCode_ As String, MinuteBar_ As String) 'MinuteBar_は"1M" , "5M"のように記述する
    
    'このインスタンスの名称
    Set mySheet = getMySheet("C_" & BrandCode_ & MinuteBar_)
    With mySheet
        .Cells(1, 1) = "=RssChart(R[1]C[0]:R[1]C[6]  , " & BrandCode_ & ",""" & MinuteBar_ & """," & ASHI & ")"  '足数（MarketSpeedの初期設定数は500）
        .Cells(2, 1) = "日付"
        .Cells(2, 2) = "時刻"
        .Cells(2, 3) = "始値"
        .Cells(2, 4) = "高値"
        .Cells(2, 5) = "安値"
        .Cells(2, 6) = "終値"
        .Cells(2, 7) = "出来高"
    End With
    mySheet.Visible = xlSheetHidden '非表示
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
            ElseIf myCandle.HighPrice > iTECBASE_MostHighCandle.HighPrice Then '同値なら新しい日時のデータのままにしておく
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

'**必要**
'格納済みのChartCandleを返す
'引数Index_を指定しないとき(Index_ = 0)は最新足を返す
'N本前の足を取得する場合は引数に「N+1」を指定すればよい (1本前なら1,24本前なら24)
Public Property Get Candle(Optional Index_ As Integer = 0) As ChartCandle
    If DicCandle.Count > Index_ Then   '4/18修正
        Set Candle = DicCandle.Items(DicCandle.Count - 1 - Index_)
    End If
End Property

'本数を表す：何本前などの算出に使用する
Private Property Get NewIndex() As Integer
    Static Counter As Integer
    Counter = Counter + 1
    NewIndex = Counter
End Property

Private Sub mySheet_Change(ByVal Target As Range)
    If isDataRange(Target) = False Then Exit Sub   'データでなければ抜ける
    
    Static PreviousClosePrice As Long
    Dim CurrentDateTime As Date
    
    '[EMA5検証用]
    '8列目に計算済みのEMA5の値を書き込みするがその書き込みは更新の対象外とする
    If Target.Item(1).Column > 7 Then Exit Sub
    
    Dim arr As Variant
    arr = Target
    '正順：古い足から格納
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        If Val(arr(i, 7)) > 0 Then   '出来高0ならば対象外
        
            CurrentDateTime = CDate(arr(i, 1) & " " & arr(i, 2)) 'Cdateでキャストすれば日付型になる（間のスペース必要）
            
            Dim myCandle As ChartCandle
            
            If Not DicCandle.Exists(CurrentDateTime) Then
                '----------------
                '新規足
                '----------------
                Set myCandle = New ChartCandle
                Call myCandle.Init(CurrentDateTime, NewIndex) 'Indexも指定
                
                'ディクショナリに格納
                Call DicCandle.Add(CurrentDateTime, myCandle)
                '規定格納数を超えたら一番過去の足を削除
'                If DicCandle.Count > MAXCOUNT Then Call DicCandle.Remove(DicCandle.Keys(0))                    '先頭を削除：ディクショナリは0オリジン
            Else
                '----------------
                '既存足
                '----------------
                Set myCandle = DicCandle.Item(CurrentDateTime)
            End If

'            Call El.SetLap
            '共通処理
            '処理時間：0.09ミリ秒程度
            'getRCI(getArrEMA5(5))をついか
            '処理時間：0.102ミリ秒程度:6/18
            
            'Arr配列の内容：{OpenPrice,HighPrice,LowPrice,CLosePrice,Volume}
            Call myCandle.addData(arr(i, 3), arr(i, 4), arr(i, 5), arr(i, 6), arr(i, 7))
            
            '6/20 SMAの計算を高速化：ミリ秒(0.033)程度
            With myCandle
                .SMA5 = getSMA5(Me)
                .SMA25 = getSMA25(Me)
                .EMA5 = getEMA5(Me)
                .EMA25 = getEMA25(Me)
            End With
'            Debug.Print El.GetLap
            
            '=============
            'イベント発火
            '=============
            RaiseEvent Change(Me)
            
            
            If PreviousClosePrice = 0 Or PreviousClosePrice <> Val(arr(i, 6)) Then
            '=============
            'イベント発火
            '=============
                RaiseEvent PriceChange(Me)
                PreviousClosePrice = Val(arr(i, 6))
            End If
        End If
    Next i

    
End Sub

