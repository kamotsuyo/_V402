VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TecBollinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'発生させるイベント
Event Change(TecBollinger_ As TecBollinger)

Public WithEvents myChart As Chart
Attribute myChart.VB_VarHelpID = -1
Public WithEvents myTick As Tick
Attribute myTick.VB_VarHelpID = -1

'このクラスインスタンスの生成元のiTECBASEクラス
Public myTecBase As iTECBASE

Private DicCandle As Dictionary
Private CollSigmaRatio As Collection

Const MAXCOUNT As Integer = 50  'DicBollingerCandleの最大格納数

'処理時間計測
Private El As ElapsedTime

Private Sub Class_Initialize()
    Set DicCandle = New Dictionary
    Set CollSigmaRatio = New Collection
    
    '処理時間計測
    Set El = New ElapsedTime
    
End Sub

'標準偏差を求めるには､次の手順に従います｡
'データの平均値を求める
'各データから平均値を引いた値（偏差）を求める
'偏差を2乗して、その和を求める
'その和をデータの総数で割って分散を求める
'分散にルートをかぶせて標準偏差を求める

Public Sub Init(Obj_ As iTECBASE)
    '引数によってインスタンス生成とAverageプロパティのセットを行う
    '1)チャートの場合
    If TypeOf Obj_ Is Chart Then
        Set myChart = Obj_
    End If
    '2)Tickの場合
    If TypeOf Obj_ Is Tick Then
        Set myTick = Obj_
    End If
    
    '生成元のTecBaseを保持しておく
    Set myTecBase = Obj_
    
    
End Sub

Public Property Get SampleCount() As Integer
    SampleCount = myTecBase.Count   '生成元のサンプル件数
End Property

'標準偏差の配列を取得する
Private Function getArrSD(N_ As Integer) As Variant
    'バグ修正
    '指定の本数分の足にディクショナリ格納数が足らないときの対応
    If DicCandle.Count < N_ Then Exit Function '戻り値はEmptyとなる
    
    Dim arr As Variant
    ReDim arr(1 To N_)
    
    Dim i As Integer
    For i = 1 To N_
        arr(i) = Candle(i - 1).StanderdDeviation
    Next i
    
    getArrSD = arr
End Function

'Sigmaの配列を取得する
Private Function getArrPrice(N_ As Integer) As Variant
    'バグ修正
    '指定の本数分の足にディクショナリ格納数が足らないときの対応
    If DicCandle.Count < N_ Then Exit Function '戻り値はEmptyとなる

    Dim arr As Variant
    ReDim arr(1 To N_)
    
    Dim i As Integer
    For i = 1 To N_
        arr(i) = Candle(i - 1).Price
    Next i

    getArrPrice = arr
End Function


'格納済みの最終TecBollingerCandleを返す
'引数Index_を指定しないときは最新足を返す
'N本前の足を取得する場合は引数に「N」を指定すればよい (1本前なら1,24本前なら24)
Public Property Get Candle(Optional Index_ As Integer = 0) As TecBollingerCandle
    If DicCandle.Count > Index_ Then
        Set Candle = DicCandle.Items(DicCandle.Count - 1 - Index_)
    End If
End Property

'標準偏差(Standerd Devitation)を返す
Private Function getStanderdDeviation(TecBase_ As iTECBASE) As Double
    Const N As Integer = 25 '母数
    If TecBase_.Count < N Then Exit Function

    Dim myCandle As iTECBASECANDLE
    Dim myAverage As Double '平均
    Dim myValue As Double '評価対象の値
    Dim myDEV As Double   '偏差(Devitation)
    Dim myTotalDEV_2 As Double      '偏差の2乗の和
    Dim myDispersion As Double '分散

    myAverage = TecBase_.Candle.SMA25   'SMA25を使う
'    myAverage = TecBase_.Candle.Ema25   'EMA25に変更してみる
    If myAverage = 0 Then Exit Function
    
    Dim i As Integer
    For i = 0 To N - 1 'この範囲指定で最新足からN個分前までを走査できる

        Set myCandle = TecBase_.Candle(i) 'Candle(i)で指定の位置の足を取得する：i=1ならば最新足
        myValue = myCandle.Price
        myDEV = myValue - myAverage '偏差：評価値 - 平均　(平均はSMA25)
'        myDEV = myValue - myCandle.SMA25   '<==この個別計算方法では数値にMarketSpeedと差が出る
        myTotalDEV_2 = myTotalDEV_2 + myDEV ^ 2 '偏差2乗の和：偏差の2乗数を合計しておく
    Next i
    
    myDispersion = myTotalDEV_2 / N '分散
    getStanderdDeviation = Sqr(myDispersion) '標準偏差：分散のルート
    
End Function

Private Function getSigmaRatioAverage(SigmaRatio_ As Double) As Double
    Const MAXCOUNT As Integer = 5
    Call CollSigmaRatio.Add(SigmaRatio_)
    If CollSigmaRatio.Count > MAXCOUNT Then Call CollSigmaRatio.Remove(1)  'コレクションは１オリジン
    
    Dim TotalSigmaRatio As Double
    
    Dim i As Integer
    For i = 1 To CollSigmaRatio.Count
        TotalSigmaRatio = TotalSigmaRatio + CollSigmaRatio(i)
    Next i
    
    getSigmaRatioAverage = TotalSigmaRatio / CollSigmaRatio.Count
End Function

Private Sub updateCandle(TecBase_ As iTECBASE)
    Dim myCandle As TecBollingerCandle
    Dim myTimeClass As String

    myTimeClass = TecBase_.Candle.TimeClass
    If DicCandle.Exists(myTimeClass) Then
        '更新
        Set myCandle = DicCandle.Item(myTimeClass)
    Else
        '新規
        Set myCandle = New TecBollingerCandle
        
        'ディクショナリへの格納
        Call DicCandle.Add(myTimeClass, myCandle)
        '規定格納数を超えたら一番過去の足を削除
        If DicCandle.Count > MAXCOUNT Then Call DicCandle.Remove(DicCandle.Keys(0))
    End If
    
'    Call El.SetLap
    '処理時間計測：ミリ秒(0.097)
    '値を格納する
    With myCandle
        .TimeClass = myTimeClass
        .Price = TecBase_.Candle.Price
        .Average = TecBase_.Candle.SMA25    'SMA25を使う
        '標準偏差などを取得して格納する
        If .Average > 0 Then
            .StanderdDeviation = getStanderdDeviation(TecBase_)
            .Sigma = .Price - .Average
        End If
        If .StanderdDeviation > 0 Then
            .SigmaRatio = .Sigma / .StanderdDeviation
            .SigmaRatioAverage = getSigmaRatioAverage(.SigmaRatio)
            .BandWidth = .StanderdDeviation * 2 / .Average * 100
        End If
        
    End With
    
'    Debug.Print TypeName(TecBase_), El.GetLap
    
    '処理時間計測：ミリ秒 (0.063)
    Call El.SetLap
    
    '標準偏差の動向を算出する
    Dim arr As Variant
'    Arr = getArrSD(5)  '5件分の標準偏差の配列に変更：6/17
    arr = getArrSD(10)  '10件分の標準偏差の配列:もとに戻した：6/18
    '指定の本数分にディクショナリ格納数がたりないときはArrはEmptyである
    If IsEmpty(arr) = False Then
        myCandle.StanderdDeviationRCI = getRCI(arr)
        myCandle.StanderdDeviationRSI = getRSI(arr)
    End If
    
   
'    Debug.Print El.GetLap, TypeName(Me), myCandle.StanderdDeviation

    RaiseEvent Change(Me)
End Sub


'生成元がChartクラスの場合のイベント受信
Private Sub myChart_Change(Chart_ As iTECBASE)
'    Call updateCandle(Chart_)
'    If Chart_.MostHighCandle Is Nothing = False Then
'        Debug.Print Chart_.MostHighCandle.HighPrice
'    End If
End Sub

Private Sub myChart_PriceChange(Chart_ As iTECBASE)
    Call updateCandle(Chart_)
End Sub

'生成元がTickクラスの場合のイベント受信
Private Sub myTick_Change(Tick_ As iTECBASE)
    Call updateCandle(Tick_)
End Sub
