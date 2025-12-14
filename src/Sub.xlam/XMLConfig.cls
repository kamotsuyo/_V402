VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public LoadStartFLG As Boolean
Public AutoMarketSpeedFLG As Boolean
Public SendMailFLG As Boolean
Public SoundFLG As Boolean

Public Sub Init()
    '読み込み対象のXMLファイルを指定する
    Dim XMLfile As String
    XMLfile = ThisWorkbook.Path & "\config.xml"
    
    'DOMDocument60を使用してXMLファイルを読み込み
    Dim DOMDoc As MSXML2.DOMDocument60
    Set DOMDoc = New MSXML2.DOMDocument60
    
    Dim Result As Boolean 'XMLファイルの読み込み結果を格納する
    Result = DOMDoc.Load(XMLfile)
    
    If Result = True Then
        'XMLファイルの読み込み成功時

        
        Dim nodeList As IXMLDOMNodeList
        '規定のXMLファイルから必要な項目を取得してプロパティにセットしていく
        
        '読み込みを開始するかどうかのフラグ
        Set nodeList = DOMDoc.getElementsByTagName("loadStart")   'タグ名は大文字小文字分別あり
        Me.LoadStartFLG = nodeList.Item(0).Text 'リストなのでitem(番号)を指定
        'マーケットスピードの再起動を自動的に行うかどうか
        Set nodeList = DOMDoc.getElementsByTagName("autoMarketSpeed")
        Me.AutoMarketSpeedFLG = nodeList.Item(0).Text
        'メール送信を行うかどうか
        Set nodeList = DOMDoc.getElementsByTagName("sendMail")
        Me.SendMailFLG = nodeList.Item(0).Text
        '音を鳴らすかどうか
        Set nodeList = DOMDoc.getElementsByTagName("sound")
        Me.SoundFLG = nodeList.Item(0).Text

        
    Else
        Call MsgBox("XMLファイル読込失敗 : " & XMLfile)
        
    End If
End Sub

