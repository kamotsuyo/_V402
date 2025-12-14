Attribute VB_Name = "TechnicalModule"
Option Explicit

'SMAを算出する
'算出数値がMarketSpeedと相違ないことを確認済：5/18
'高速化を目的に、計算方式を変更：6/20
'もとに戻した：数値の相違が大きい
Public Function getSMA5(TecBase_ As iTECBASE) As Double
    Const N As Integer = 5
    'DicTickCandleを逆順で走査
    Dim Total As Long
    Select Case TecBase_.Count
        Case Is < N
            Exit Function
        Case Is >= N
            Dim i As Integer
            For i = 0 To N - 1     'TecBase_.Candle(0)は最新足
                Total = Total + TecBase_.Candle(i).Price
            Next i
            getSMA5 = Total / N
    End Select
    
End Function

Public Function getSMA25(TecBase_ As iTECBASE) As Double
    Const N As Integer = 25
    'DicTickCandleを逆順で走査

    Dim Total As Long
    Select Case TecBase_.Count
        Case Is < N
            Exit Function
        Case Is >= N
            Dim i As Integer
            For i = 0 To N - 1     'TecBase_.Candle(0)は最新足
                Total = Total + TecBase_.Candle(i).Price
            Next i
            getSMA25 = Total / N
'        Case Is > N
'            getSMA25 = (TecBase_.Candle(1).SMA25 * (N - 1) + TecBase_.Candle.Price) / N
    End Select
    
End Function

Public Function getSMAvolume(TecBase_ As iTECBASE, Optional N_ As Integer = 5) As Double
    'DicTickCandleを逆順で走査
    Dim Total As Long
    Select Case TecBase_.Count
        Case Is < N_
            Exit Function
        Case Is >= N_
            Dim i As Integer
            For i = 0 To N_ - 1     'TecBase_.Candle(0)は最新足
                Total = Total + TecBase_.Candle(i).Volume
            Next i
            getSMAvolume = Total / N_
    End Select
    
End Function

'EMA5を算出する
'算出数値がMarketSpeedと相違ないことを確認済：5/18
Public Function getEMA5(TecBase_ As iTECBASE) As Double
    '1日目の算出式（Term=5の場合、チャートの5本目の値を算出）
    '指定期間の直近n日の移動平均とする。 移動平均の算出は「移動平均線」参照
    '2日目以降の算出式（Term=5の場合、チャートの6本目以降の値を算出）
    'EMAt   = EMA1 + (2/(Term+1)) * (Ct-EMA1)
    
    Const N As Integer = 5

    'EMA5について
    If TecBase_.Count = N Then
        getEMA5 = getSMA5(TecBase_)
    ElseIf TecBase_.Count > N Then
        getEMA5 = (TecBase_.Candle(1).EMA5 * (N - 1) + TecBase_.Candle.Price * 2) / (N + 1) '1本前の足はMe.Candle(1)
    End If
    
End Function
'EMA25を算出する
'算出数値がMarketSpeedと相違ないことを確認済：5/18
Public Function getEMA25(TecBase_ As iTECBASE) As Double
    '1日目の算出式（Term=5の場合、チャートの5本目の値を算出）
    '指定期間の直近n日の移動平均とする。 移動平均の算出は「移動平均線」参照
    '2日目以降の算出式（Term=5の場合、チャートの6本目以降の値を算出）
    'EMAt   = EMA1 + (2/(Term+1)) * (Ct-EMA1)
    
    Const N As Integer = 25

    'EMA25について
    If TecBase_.Count = N Then
        getEMA25 = getSMA25(TecBase_)
    ElseIf TecBase_.Count > N Then
        getEMA25 = (TecBase_.Candle(1).EMA25 * (N - 1) + TecBase_.Candle.Price * 2) / (N + 1) '1本前の足はMe.Candle(1)
    End If
    
End Function
'++++++++++++++++++++++
'価格データの中に同じ値が含まれる場合、その順位を平均して算出する方法が一般的です。
'例えば､以下のように2日目と3日目の終値がともに102円の場合､2日目と3日目の順位には優劣をつけず､平均値を取って求められます｡
'--->
'7/11修正：[価格順位を平均]しなくてはならない。価格順位は i である(ArrTimeIndex(i-1) は　時間順位なので誤り！)
'++++++++++++++++++++++
Public Function getRCI(ByVal Arr_ As Variant) As Double
    On Error GoTo ErrorHandler
    
    If IsEmpty(Arr_) Then Exit Function
    
    '経過秒数
'    Dim El As ElapsedTime
'    Set El = NewElapsedTime

    Dim N As Integer    '標本数
    N = UBound(Arr_) - LBound(Arr_) + 1
    
    Dim D As Double     '日付の順位と価格の差を2乗し、合計した数値

    Dim i As Integer
    Dim j As Integer
    
    Dim ArrSourceData As Variant     '標本の価格を格納する配列
    Dim ArrTimeIndex As Variant     '時刻順位を管理する配列
    Dim ArrHighIndex As Variant     '補正した価格順位を格納する配列
    ReDim ArrSourceData(1 To N)
    ReDim ArrTimeIndex(1 To N)
    ReDim ArrHighIndex(1 To N)
    
    Dim Index As Integer    '順序となる値
    
    'ArrPとArrTを生成
    'ArrPも基数１から始めたいのであえてループ処理でセットする
    For i = LBound(Arr_) To UBound(Arr_)
        Index = Index + 1   '順序番号：１から始める
        ArrTimeIndex(Index) = Index
        ArrSourceData(Index) = Arr_(i)
    Next i
    

'    Debug.Print "---------------------------", N
'    Debug.Print Now; "開始時"; Join(ArrTimeIndex), Join(ArrSourceData)
           
    '与えられた配列をもとに正順にソートする
    Dim Temp As Variant '一時退避用
    For i = LBound(ArrSourceData) To UBound(ArrSourceData) - 1    '1つ少なくループ：これは配列ソートの手順
        For j = i + 1 To UBound(ArrSourceData)
            If ArrSourceData(i) < ArrSourceData(j) Then   '大きい順に並び替え
                '与えられた配列をソートする
                Temp = ArrSourceData(j)
                ArrSourceData(j) = ArrSourceData(i)
                ArrSourceData(i) = Temp
                '同時に時刻配列もソートする
                Temp = ArrTimeIndex(j)
                ArrTimeIndex(j) = ArrTimeIndex(i)
                ArrTimeIndex(i) = Temp
            End If
        Next j
    Next i
    
    '価格順位の配列ArrHをセットする???
    For i = 1 To N
        ArrHighIndex(i) = i
    Next i
    
'    Debug.Print "Sort済"; Join(ArrTimeIndex), Join(ArrSourceData), Join(ArrHighIndex)

    '+++++++++++++++++++++++
    '補正部分
    '+++++++++++++++++++++++
    'ソート済の配列を使って同一価格の補正を行う
    '価格データの中に同じ値が含まれる場合､その[価格順位]を平均して算出する方法が一般的です｡
    '参照：https://www.fintokei.com/jp/blog/rci-calculation/

'    Call El.SetLap
    
    '所要時間　約0.15ミリ秒程度
    '->ディクショナリを使用するより高速なのでこちらを使う
    Dim DicSame As Dictionary
    If DicSame Is Nothing Then Set DicSame = New Dictionary
    Dim SameCounter As Integer
    Dim SumHighIndex As Integer
    For i = 1 To N - 1
        Dim CurrentIndex As Integer
        Dim NextIndex As Integer
        CurrentIndex = i
        NextIndex = i + 1
        If ArrSourceData(CurrentIndex) = ArrSourceData(NextIndex) Then
            Select Case SameCounter
                Case 0
                    SameCounter = 1
                    SumHighIndex = CurrentIndex + NextIndex
                Case Is >= 1
                    SameCounter = SameCounter + 1
                    SumHighIndex = SumHighIndex + NextIndex
            End Select

            For j = SameCounter To 0 Step -1
'                Debug.Print j, i - j, SumHighIndex / (SameCounter + 1)
                ArrHighIndex(i - j + 1) = SumHighIndex / (SameCounter + 1)
            Next j
        Else
            SameCounter = 0
            SumHighIndex = 0
        End If

    Next i

'    Debug.Print El.GetLap, "補正後", Join(ArrTimeIndex), Join(ArrSourceData), Join(ArrHighIndex)
    

'    D = 0
    'RCI計算
    For i = 1 To N
        D = D + (ArrHighIndex(i) - ArrTimeIndex(i)) ^ 2
'        Debug.Print i; ArrHighIndex(i); (ArrHighIndex(i) - i) ^ 2; D
    Next i

    getRCI = (1 - (6 * D) / (N ^ 3 - N)) * 100
'    Debug.Print "補正 RCI"; Format(getRCI, "0")
Exit Function
ErrorHandler:
    getRCI = 0

End Function


'RSI（Relative Strength）の計算
'平均上昇幅と平均下降幅の計算: 一定期間（例: 14日間）における価格上昇幅の平均を計算します。
'同様に､一定期間における価格下降幅の平均を計算します｡
'上昇幅・下降幅は、前日の終値との差分で計算します。
'**同一価格の場合：が続く場合は､上昇幅も下降幅も0として計算に含めます｡（つまりなにもしない）

'2025/7/8：この計算式で問題なし。今まで、マーケットスピードの数値と差異が発生していた原因は
'サンプル件数の差である。RSIの期間１０であれば、差を１０件分計算するので、サンプル件数は１１件必要である。
Public Function getRSI(Arr_ As Variant) As Double
    If IsEmpty(Arr_) Then Exit Function
    
    Dim N As Integer    '期間N
    N = UBound(Arr_) - LBound(Arr_) + 1
    
'    Debug.Print "--", N, Join(Arr_)
    
    Dim UpSum As Double   '上昇分合計
    Dim DownSum As Double '下落分合計
    
    Dim i As Integer
    For i = LBound(Arr_) To UBound(Arr_) - 1 '最新の時刻データから走査
        
        Dim Fluctuation As Double 'データ間の増減
        Fluctuation = Arr_(i) - Arr_(i + 1)   '１つ過去のデータとの差(直近が増加していれば+になる)
        
'        Debug.Print i, Fluctuation, Arr_(i), Arr_(i + 1)
        
        Select Case Fluctuation
            Case Is > 0
                UpSum = UpSum + Abs(Fluctuation)
            Case Is <= 0
                DownSum = DownSum + Abs(Fluctuation)
            Case 0
                'なにもしない
        End Select
    Next i
    
    
    
    If UpSum + DownSum > 0 Then
        getRSI = UpSum / (UpSum + DownSum) * 100
    End If
    
End Function


'Priceの配列を取得する
'第3引数のStart_は開始点：0ならば現在足から
Public Function getArrPrice(TecBase_ As iTECBASE, N_ As Integer) As Variant
    On Error GoTo ErrorHandler
    
    'バグ修正
    '指定の本数分の足にディクショナリ格納数が足らないときの対応
    '--->ErrorHandlerで対応
'    If TecBase_.Count < 5 + N_ Then Exit Function  '戻り値はEmptyとなる

    Dim Arr As Variant '戻り値となる配列
    ReDim Arr(1 To N_)  '１から始めたい：getRCIメソッドへ引数として渡すため。
    
    Dim myCandle As iTECBASECANDLE
    
    Dim i As Integer
    For i = 1 To N_
        '戻り値の配列を1から始まる。よってiが1から始まるので (i-1) とする

        Set myCandle = TecBase_.Candle(i - 1)    'Candle(0)は現在足
        Arr(i) = myCandle.Price '最新の時刻データから順次配列に格納する

    Next i
    
    getArrPrice = Arr
Exit Function
ErrorHandler:
    Debug.Print "getArrPrice Error", Err, N_
    getArrPrice = Empty
End Function

'Volumeの配列を取得する
Public Function getArrVolume(TecBase_ As iTECBASE, N_ As Integer) As Variant
    'バグ修正
    '指定の本数分の足にディクショナリ格納数が足らないときの対応
    If TecBase_.Count < 5 + N_ Then Exit Function  '戻り値はEmptyとなる

    Dim Arr As Variant '戻り値となる配列
    ReDim Arr(1 To N_)  '１から始めたい：getRCIメソッドへ引数として渡すため。
    
    Dim myCandle As iTECBASECANDLE
    
    Dim i As Integer
    For i = 1 To N_
        '戻り値の配列を1から始まる。よってiが1から始まるので (i-1) とする
        Set myCandle = TecBase_.Candle(i - 1)   'Candle(0)は現在足
        Arr(i) = myCandle.Volume '最新の時刻データから順次配列に格納する
    Next i
    
    getArrVolume = Arr
End Function

'SMA5_Volumeの配列を取得する
Public Function getArrEMA5(TecBase_ As iTECBASE, N_ As Integer) As Variant
    'バグ修正
    '指定の本数分の足にディクショナリ格納数が足らないときの対応
    If TecBase_.Count < 5 + N_ Then Exit Function '戻り値はEmptyとなる

    Dim Arr As Variant '戻り値となる配列
    ReDim Arr(1 To N_)  '１から始めたい：getRCIメソッドへ引数として渡すため。
    
    Dim myCandle As iTECBASECANDLE
    
    Dim i As Integer
    For i = 1 To N_
        '戻り値の配列を1から始まる。よってiが1から始まるので (i-1) とする
        Set myCandle = TecBase_.Candle(i - 1)   'Candle(0)は現在足
        Arr(i) = myCandle.EMA5 '最新の時刻データから順次配列に格納する
    Next i

    getArrEMA5 = Arr
End Function

