VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StraHandTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'###########################
Implements iSTRA
'###########################

'======================
'STRA共通

'Initにてセットするプロパティ
Private myBrand As iBRAND
Private myDeayType As En_DealType

Private OrderLock As Boolean    '必要
Private CurrentProfitLoss As Long '必要
Private TotalProfitLoss As Long '必要
'======================

'フォーム
Private myForm As F_StraHandTest

Private WithEvents Button1 As CommandButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2 As CommandButton
Attribute Button2.VB_VarHelpID = -1

'----------------
'Test
'Private myProp As ty_PropStra
'内部プロパティ
'Private Property Get isOutSideHours() As Boolean
''営業時間外 発注対象でない時刻帯を記述する
'    Dim Result As Boolean   '初期値はFalse
'    Select Case TimeValue(Now)
'        'TimeValue(Now) はDate型で戻り値は 7:15:45 のようになるが範囲指定判定では秒は省略される
'        '例：TimeValue("06:00") To TimeValue("08:45")　の場合
'        '06:00:00 から 08:45:00までがtrueと判定される。(08:45:01以降はFalseである)
'        '指定文字列は秒数まで指定すればよい(右側だけでもよい)
'        'TimeValue("06:00") To TimeValue("08:45:59") ==>　これなら 08:45台はすべてTrue判定になる
'
''        Case TimeValue("06:00") To TimeValue("08:45") '開始前
''            Result = True
''        Case TimeValue("08:45") To TimeValue("09:00") '前場準備
''            Result = True
''        Case TimeValue("15:30") To TimeValue("15:45") '前場おわり際
''            Result = True
''        Case TimeValue("17:00") To TimeValue("18:00") '夜間取引の閑散時間帯
''            Result = True
''        Case TimeValue("05:30") To TimeValue("06:00") '夜間終わり際
''            Result = True
'    End Select
'    isOutSideHours = Result
'End Property

Private Sub Class_Initialize()
    'フォーム：複製で開く
    Set myForm = UserForms.Add("F_StraHandTest")
    myForm.Show vbModeless
    
    Set Button1 = myForm.CommandButton1
    Set Button2 = myForm.CommandButton2

End Sub






'=============Stra共通=============

Private Sub iSTRA_Init(Brand_ As SubProject.iBRAND, DealType_ As SubProject.En_DealType)
    Set myBrand = Brand_
    myDeayType = DealType_
    
    myForm.Caption = iSTRA_Name 'フォーム名称
    
    'CollStraコレクションに追加
    Call CollStra.Add(Me) 'コレクションに追加
    
End Sub
Private Sub iSTRA_changeOrderLock(Flg_ As Boolean)
    OrderLock = Flg_
End Sub
Private Sub iSTRA_Callback(Order_ As iORDER)
    '個別ストラテジの合計損益を更新
    TotalProfitLoss = TotalProfitLoss + Order_.ProfitLoss

    'フォームを更新する
    Call F_Strategys.Display(Me)
End Sub

Private Property Get iSTRA_isOutSideHours() As Boolean
'営業時間外 発注対象でない時刻帯を記述する
    Dim Result As Boolean   '初期値はFalse
    Select Case TimeValue(Now)
        'TimeValue(Now) はDate型で戻り値は 7:15:45 のようになるが範囲指定判定では秒は省略される
        '例：TimeValue("06:00") To TimeValue("08:45")　の場合
        '06:00:00 から 08:45:00までがtrueと判定される。(08:45:01以降はFalseである)
        '指定文字列は秒数まで指定すればよい(右側だけでもよい)
        'TimeValue("06:00") To TimeValue("08:45:59") ==>　これなら 08:45台はすべてTrue判定になる
        
'        Case TimeValue("06:00") To TimeValue("08:45") '開始前
'            Result = True
'        Case TimeValue("08:45") To TimeValue("09:00") '前場準備
'            Result = True
'        Case TimeValue("15:30") To TimeValue("15:45") '前場おわり際
'            Result = True
'        Case TimeValue("17:00") To TimeValue("18:00") '夜間取引の閑散時間帯
'            Result = True
'        Case TimeValue("05:30") To TimeValue("06:00") '夜間終わり際
'            Result = True
    End Select
    iSTRA_isOutSideHours = Result
End Property
Private Property Get iSTRA_DealType() As SubProject.En_DealType
    iSTRA_DealType = myDeayType
End Property
Private Property Get iSTRA_MAXORDERQUANTITY() As Integer
    iSTRA_MAXORDERQUANTITY = 2
End Property
Private Property Get iSTRA_Name() As String
    iSTRA_Name = TypeName(Me) & ":" & myBrand.PetName & ":" & myDeayType
End Property
Private Property Get iSTRA_AutoExit() As SubProject.iAUTOEXIT
    'なにも記述しなければNothingになる
'    Set iSTRA_AutoExit = New iAUTOEXITType1
End Property
Private Property Get iSTRA_OrderLock() As Boolean
    iSTRA_OrderLock = OrderLock
End Property
Private Property Get iSTRA_TotalCount() As Long
    iSTRA_TotalCount = getStraTotalCount(Me)
End Property
Private Property Get iSTRA_AliveCount() As Long
    iSTRA_AliveCount = getStraAliveCount(Me)
End Property
Private Property Get iSTRA_Brand() As SubProject.iBRAND
    Set iSTRA_Brand = myBrand
End Property
Private Property Get iSTRA_StanderdMargin() As Long
    iSTRA_StanderdMargin = 30 '<---ここに設定値を記述する
End Property
Private Property Get iSTRA_StanderdLoss() As Long
    iSTRA_StanderdLoss = 30    '<---ここに設定値を記述する(Loss +数値で書く)
End Property
Private Property Get iSTRA_CurrentProfitLoss() As Long
    iSTRA_CurrentProfitLoss = CurrentProfitLoss
End Property
Private Property Get iSTRA_TotalProfitLoss() As Long
    iSTRA_TotalProfitLoss = TotalProfitLoss
End Property
'==========================================

Private Sub Button1_Click()
    '売
    '価格をマーケットの買気配値にする：成り行き注文と同じ
    Dim myOrder As iORDER
    Set myOrder = getNewOrder(Me, 1)
    If Not myOrder Is Nothing Then
        Call myOrder.PublishOpen(売, myBrand.BuyPrice) '買気配値で売る：成り行き注文と同じ
    End If
End Sub
Private Sub Button2_Click()
    '買
    Dim myOrder As iORDER
    Set myOrder = getNewOrder(Me, 1)
    If Not myOrder Is Nothing Then
        Call myOrder.PublishOpen(買, myBrand.SellPrice) '売気配値で買う：成り行き注文と同じ
    End If
End Sub



