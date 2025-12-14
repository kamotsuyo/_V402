VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_Stra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StraTest As iSTRA
Private StraTest2 As iSTRA
Private StraHand As iSTRA

Private Sub Class_Initialize()
    Set StraTest = New StraHandTest
    Call StraTest.Init(MMBrand.MN225MF, DummyOrder)
    
    Set StraTest2 = New StraHandTest
    Call StraTest2.Init(MMBrand.M8798, DummyOrder) '動作確認OK
    
    Set StraHand = New StraHandTest
    Call StraHand.Init(MMBrand.MN225Micro, FutureOrder)
    '*CollStraコレクション追加はStra.Init内で行われる
    
    '------------------
    'F_Strategysフォーム初期セット
    '------------------
    Call F_Strategys.Init

End Sub

