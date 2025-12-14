VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MM_SubManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event ConnctChange(Status_ As String)
Event CodeResolved()

Private myXmlConf As XMLConfig
Private WithEvents MMConn As MM_Connection
Attribute MMConn.VB_VarHelpID = -1
'公開プロパティ
Public WithEvents MMCode As MM_FutureCode
Attribute MMCode.VB_VarHelpID = -1

Private Sub Class_Initialize()

    '-----------------
    '設定ファイルXMLConfigの読み込み
    '-----------------
    Set myXmlConf = New XMLConfig
    Call myXmlConf.Init
    
    If myXmlConf.LoadStartFLG Then
        FormSub.addListItem "XMLConfigロード完了"

        '-------------------
        '公開用プロパティのインスタンス生成
        '-------------------
        Set MMConn = New MM_Connection
        
    Else
        FormSub.addListItem "loadStartFLGがFalse"
    
    End If
End Sub

Private Sub Class_Terminate()
    Set MMConn = Nothing
    Set MMCode = Nothing
End Sub

Private Sub MMCode_AllResolved()
    RaiseEvent CodeResolved
End Sub

Private Sub MMConn_ReConnect()

End Sub

Private Sub MMConn_RssConnectChange(Status_ As String)
    Debug.Print "MMConn_RssConnectChange", Status_
    
    RaiseEvent ConnctChange(Status_)
    
    If Status_ = "完了" Then
        '先物コード解決を実行する
        If MMCode Is Nothing Then
            Set MMCode = New MM_FutureCode
        End If

    
    End If
End Sub
