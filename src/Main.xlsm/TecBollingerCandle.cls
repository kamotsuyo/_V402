VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TecBollingerCandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TimeClass As String    '評価対象時間階級
Public Average As Double '平均
Public Price As Double
Public StanderdDeviation As Double  '標準偏差
Public Sigma As Double   '標本偏差
Public SigmaRatio As Double        '標本偏差/標準偏差
'7/9追加
Public BandWidth As Double      '標準偏差/平均

'計算結果を格納する
Public SigmaRatioAverage As Double
Public StanderdDeviationRSI As Double
Public StanderdDeviationRCI As Double
