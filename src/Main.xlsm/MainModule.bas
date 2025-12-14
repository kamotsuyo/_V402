Attribute VB_Name = "MainModule"
Option Explicit

Public CollStra As Collection
Public DickOorder As Dictionary
Public DicPosition As Dictionary

Public MMRaktenList As MM_RaktenList
Public MMBrand As MM_Brand
Public MMTec As MM_Tec
Public MMStra As MM_Stra

'------------------
'プログラム開始
'------------------
Public Sub Start()
    Set CollStra = New Collection
    Set DickOorder = New Dictionary
    Set DicPosition = New Dictionary

    '初期化
    Set MMRaktenList = Nothing
    Set MMBrand = Nothing
    Set MMTec = Nothing
    Set MMStra = Nothing
    
    
    '!!以下のプロパティ宣言はMainModuleにて宣言済
    Set MMBrand = New MM_Brand
    Set MMRaktenList = New MM_RaktenList
    Set MMTec = New MM_Tec
    Set MMStra = New MM_Stra
    
    
    'SendKey実行
    'MarketSpeed発注不可==>発注可切り替え
    Call SendKey_MarketSpeedOrderSwitch
    

    
End Sub
