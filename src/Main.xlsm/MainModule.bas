Attribute VB_Name = "MainModule"
Option Explicit

Public CollStra As Collection
Public DickOorder As Dictionary
Public DicPosition As Dictionary

Public MMBrand As MM_Brand
Public MMTec As MM_Tec
Public MMStra As MM_Stra

'シングルトンクラス
Public Property Get MMList() As MM_RaktenList
    Dim Singleton As MM_RaktenList
    If Singleton Is Nothing Then
        Set Singleton = New MM_RaktenList
    End If
    Set MMList = Singleton
End Property

'------------------
'プログラム開始
'------------------
Public Sub Start()

    Set CollStra = New Collection
    Set DickOorder = New Dictionary
    Set DicPosition = New Dictionary

    '初期化
    Set MMBrand = Nothing
    Set MMTec = Nothing
    Set MMStra = Nothing
    
    
    '!!以下のプロパティ宣言はMainModuleにて宣言済
    Set MMBrand = New MM_Brand
    Set MMTec = New MM_Tec
    Set MMStra = New MM_Stra
    
    
    'SendKey実行
    'MarketSpeed発注不可==>発注可切り替え
    Call SendKey_MarketSpeedOrderSwitch
    

    
End Sub
