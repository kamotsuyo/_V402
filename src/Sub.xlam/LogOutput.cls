VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iASYNCGATE
'#####################

Private LogFile As String   '書き込みのファイル名称
Private FileNum As Integer  'Openメソッドで使用するFileNumを宣言

Public Sub Init(DirName_ As String)

    Dim AppDir As String
    Dim SaveDir As String
    Dim FileName As String
    
    'ディレクトリ作成
    AppDir = ThisWorkbook.Path 'このアプリのディレクトリ
    SaveDir = AppDir & "\" & DirName_   '引数からセットする
    
    'ファイル名称作成 日付でファイル名称を作る

    FileName = Format(Date, "yyyymmdd") & ".log"

    
    'ログファイル名称　統合
    LogFile = SaveDir & "\" & FileName '拡張子をtxtからlogに変更。（.gitignoreの対象外）
    'openメソッドで使用するFileNum
    FileNum = FreeFile 'FreeFile関数は使用可能なファイル番号を返します。
    
    'ディレクトリの存在確認：なければMkDirで作成
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir (SaveDir)
    End If
    
   
End Sub

'引数Log_は配列
Public Sub AddLog(ArrLog_ As Variant)
    '非同期で呼び出し
    Call AsyncModule.setStack(Me, ArrLog_)

End Sub

Private Sub iASYNCGATE_GATE(Args_ As Variant)
    Dim Str As String
    Select Case VarType(Args_)
        Case vbVariant + vbArray
            Str = Join(Args_, ",")
        Case Else
            Debug.Print "printLogの対象ではありません" & VarType(Args_)
            Exit Sub
    End Select
    
    Open LogFile For Append Shared As #FileNum

    Print #FileNum, Str

    Close #FileNum
End Sub
