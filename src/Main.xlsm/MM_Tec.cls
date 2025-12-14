VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_Tec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------
'テクニカル
'-----------------

'基本的テクニカル
Public WithEvents Chart_N225MF_1M As Chart
Attribute Chart_N225MF_1M.VB_VarHelpID = -1
Public WithEvents Tick_N225MF_15S As Tick
Attribute Tick_N225MF_15S.VB_VarHelpID = -1

'発展型テクニカル：基本テクニカルの後に起動する
Public WithEvents Bollinger_N225MF_1M As TecBollinger
Attribute Bollinger_N225MF_1M.VB_VarHelpID = -1
Public WithEvents Bollinger_N225MF_15S As TecBollinger
Attribute Bollinger_N225MF_15S.VB_VarHelpID = -1

Implements iASYNCGATE

Private Sub Class_Initialize()
    
    'ここからsetDelayStack可能
'    Call AsyncModule.setDelayStack(Me, TimeValue("00:00:03"))
    
    Set Chart_N225MF_1M = New Chart
    Call Chart_N225MF_1M.Init(SubManager.MMCode.N225MF.BrandCode, "1M")
    
    Set Tick_N225MF_15S = New Tick
    Call Tick_N225MF_15S.Init(SubManager.MMCode.N225MF.BrandCode, "15")
    

    
    '---------------
    'テクニカル
    '---------------
    Set Bollinger_N225MF_1M = New TecBollinger
    Call Bollinger_N225MF_1M.Init(Chart_N225MF_1M)
    
    Set Bollinger_N225MF_15S = New TecBollinger
    Call Bollinger_N225MF_15S.Init(Tick_N225MF_15S)
End Sub


Private Sub iASYNCGATE_GATE(Args_ As Variant)
    Debug.Print "hoge"
End Sub

