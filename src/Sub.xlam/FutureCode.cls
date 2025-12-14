VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FutureCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private myMother As MM_FutureCode

Public PetName As String    'FutureCodeResolverで参照するのでPublic
Public BrandCode As String  '決定した先物コード

'以下、確認用に保持
Private Expantion As String
Private BrandName As String

Private Sub Class_Initialize()

End Sub

'先物の「銘柄種類」:PetNameは以下の文字列
'225先物：N225F
'225mini : N225MF
'225マイクロ : N225UF
Public Sub Init(PetName_ As String, Mother_ As MM_FutureCode)

    PetName = PetName_
    Set myMother = Mother_
    
    
    '先物コード解決
    Dim myResolver As FutureCodeResolver
    Set myResolver = New FutureCodeResolver
    Call myResolver.Init(Me)

End Sub

'FutureCodeResolverインスタンスからCallされる
Public Sub Resolved(Resolver_ As FutureCodeResolver)
    BrandCode = Resolver_.BrandCode
    '以下、確認テスト用に保持しておく
    Expantion = Resolver_.Expantion
    BrandName = Resolver_.BrandName

   
    Call myMother.CheckResolvedAll
        
End Sub
