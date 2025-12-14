VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElapsedTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*************************
'経過時間（ミリ秒）を算出するだけのクラス
'1秒 = 1000ミリ秒
'-----------------------------
'2025/3/18　平均値機能追加
'使用方法：
'1)一度だけ計測でよければ
'宣言：Dim el as ElapedTime
'インスタンス生成 : Set el = new ElapedTime
'--計測対象の処理記述--
'el.getLap() で計測したミリ秒文字列取得
'2)繰り返し計測して平均をみたいとき
'宣言とインスタンス生成を別処理で行う。(インスタンスを使いまわす)
'el.setLap() で計測区間開始
'--計測対象の処理記述--
'el.getLap() で計測したミリ秒文字列取得
'となる
'-----------------------------

Private StartMicroSecond As String

'3/18追記：平均を算出するため

Private Sub Class_Initialize()
    StartMicroSecond = getMicroSecond
End Sub

Public Sub SetLap()
    StartMicroSecond = getMicroSecond
End Sub

Public Function GetLap() As String
    Static Counter As Integer
    Static LapTotal As Double
    Dim LapAverage As Double
    Counter = Counter + 1
    
    GetLap = getMicroSecond - StartMicroSecond  'この時点で計測終わり。あとの処理時間は対象外
    
    '平均値を算出する
    LapTotal = LapTotal + GetLap
    LapAverage = LapTotal / Counter

    '結果はミリ秒単位で返す : ミリ秒 = マイクロ秒 *1000
    GetLap = Format(GetLap * 1000, "0.000 ミリ秒") & "(" & Format(LapAverage * 1000, "0.000") & ")"

    
End Function




