VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_FutureCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'コレクション：Privateなのでmyを付けておく（分別）
Private myCollFutureCodes As Collection    '生成した先物インスタンスを格納

'--------------------
'先物の「銘柄種類」は以下の文字列
'225先物：N225F
'225mini : N225MF
'225マイクロ : N225UF
Public N225F As FutureCode
Attribute N225F.VB_VarHelpID = -1
Public N225MF As FutureCode
Attribute N225MF.VB_VarHelpID = -1
Public N225micro As FutureCode
Attribute N225micro.VB_VarHelpID = -1

'全先物コード解決イベント
Event AllResolved()

Private Sub Class_Initialize()
    '先物コレクションへのセット
    If myCollFutureCodes Is Nothing Then Set myCollFutureCodes = New Collection
    
    '先物インスタンス生成開始
    '以下の順番に処理をする
    '各インスタンス生成 => コレクションに格納 => Initをコール
    Set N225F = New FutureCode '1)225先物
    Set N225MF = New FutureCode '2)225先物ミニ
    Set N225micro = New FutureCode '3)225先物マイクロ
    
    '**各インスタンス生成してからコレクションに格納する
    Call myCollFutureCodes.Add(N225F)
    Call myCollFutureCodes.Add(N225MF)
    Call myCollFutureCodes.Add(N225micro)
    
    '**Initは最後にまとめて行う
    Call N225F.Init("N225F", Me)
    Call N225MF.Init("N225MF", Me)
    Call N225micro.Init("N225UF", Me)
End Sub


Public Sub CheckResolvedAll()
    'このフォームインスタンスのプロパティにアクセスできるが、格納した内容にアクセスできない(Nothingとなる)
    'myCollFutureCodesコレクションをこのインスタンスでなくModuleに記述したら解決した
    Dim myFutureCode As FutureCode
    For Each myFutureCode In myCollFutureCodes
        If myFutureCode.BrandCode = "" Then Exit Sub
    Next
    
    'フォームへログ
    FormSub.addListItem ("全先物のコード取得完了済")
    Debug.Print "全先物のコード取得完了済"
    
    '全先物のコード取得完了済
    '----------------------
    RaiseEvent AllResolved
    

End Sub

