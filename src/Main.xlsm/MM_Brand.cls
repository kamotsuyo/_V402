VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_Brand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************
'FinancialProdut起動用
'Futureクラスインスタンス
'Stockクラスインスタンス


'-----------------
'商品
'-----------------
'先物マーケット情報
Public MN225MF As Future
Public MN225Micro As Future

'株式マーケット情報
Public M8798 As Stock
Public M9983 As Stock 'ファーストリテイリング
Public M9432 As Stock 'NTT
Public M9343 As Stock 'ソフトバンク

Private Sub Class_Initialize()

    'マーケット
    '先物：先物の場合はFutureCodeクラスを引数にセットする：コードではなくクラスそのまま
    Set MN225MF = New Future
    Call MN225MF.Init(SubManager.MMCode.N225MF)
    Set MN225Micro = New Future
    Call MN225Micro.Init(SubManager.MMCode.N225micro)
        
    '株式
    Set M8798 = New Stock
    Call M8798.Init("8798")
    
    Set M9983 = New Stock
    Call M9983.Init("9983")
    
    Set M9432 = New Stock
    Call M9432.Init("9432")
    
    Set M9343 = New Stock
    Call M9343.Init("9343")
End Sub
