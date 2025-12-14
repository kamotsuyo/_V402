VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FutureCodeResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#####################
Implements iASYNCGATE
'#####################

Private mySheet As Worksheet
Attribute mySheet.VB_VarHelpID = -1

Private myFutureCode As FutureCode
Private ArrGengetsuLength As Integer  '限月の候補数：１件なら１。2件なら２

'取得した先物コードなど
Public BrandName As String
Public BrandCode As String
Public Expantion As String

Public Sub Init(FutureCode_ As FutureCode)
    
    Set mySheet = ThisWorkbook.Worksheets.Add   '名前をつけずにシートを追加作成
    
    '引数のFutureCodeをプロパティにセット
    Set myFutureCode = FutureCode_
    
    ArrGengetsuLength = UBound(ArrGengetsu) + 1 'ArrGengetsuは０オリジン
    
    Dim i As Integer
    For i = 0 To UBound(ArrGengetsu)
        With mySheet
            .Cells(i + 1, 1) = "=RssFOPCode(""" & myFutureCode.PetName & """, " & ArrGengetsu(i) & ")"
            .Cells(i + 1, 2) = "=RssFOPMarket(R[0]C[-1], " & """残存日数""" & ")"
            .Cells(i + 1, 3) = "=RssFOPMarket(R[0]C[-2], " & """銘柄名称""" & ")"
        End With
    Next i
    
    '-------------------------
    '待機
    '-------------------------
    'mySheetの更新までの一定時間を待って処理する
'    Call AsyncModule.setDelayStack(Me, TimeValue("00:00:05"))   '5秒：
'    Call AsyncModule.setDelayStack(Me, TimeValue("00:00:02"))   '2秒：単独なら処理OKだが、Main起動時には不安定
'    Call AsyncModule.setDelayStack(Me, TimeValue("00:00:03"))   '3秒：処理OK
    Call AsyncModule.setDelayStack(Me, TimeValue("00:00:02"))   '2秒
End Sub

'先物コードの取得の際の候補限月を配列で返す
Private Property Get ArrGengetsu() As Variant
'-----------------------------------------
'中心限月に関する問題点と対策（2024/3/4追記）
'現在の計算方式では6月限が中心限月になるが、MarketSpeedの中心限月は3月限である。
'これは中心限月の切り替えが「SQ決済日」として証券取引所にて規定されており、今回のケースでいえば
'日経225miniの3月限の最終取引日は3/7(木曜)となっている。それまでは3月限が中心限月なのだ。
'SQ決済日は、その限月の該当する月の第2金曜日であるのでその前営業日が最終取引日となる
'まとめると
'限月（日経225）は3月,6月,9月,12月と4回存在する(日経225miniは毎月の限月があるが、取引が多いのは左記の4回)
'したがってminiも取引対象の限月は3月,6月,9月,12月とする。
'現在の月と中心限月の関係は以下の通り
'1月:3月限が中心限月
'2月:3月限が中心限月
'3月:3月限あるいは6月限が中心限月
'4月:6月限が中心限月
'5月:6月限が中心限月
'6月:6月限あるいは9月限が中心限月
'7月:9月限が中心限月
'8月:9月限が中心限月
'9月:9月限あるいは12月限が中心限月
'10月:12月限が中心限月
'11月:12月限が中心限月
'12月:12月限あるいは翌年3月限が中心限月
    Dim Arr As Variant
    Dim myYear As Integer
    Dim myMonth As Integer
    myYear = Year(Now)
    myMonth = Month(Now)
    'test
'    myMonth = 12
    Select Case myMonth  '現在の月
        Case 1, 2
            ReDim Arr(0)
            Arr(0) = myYear & "03" '0埋め :3月なら03=>限月は202303など6桁にしなくてはならない
        Case 3
            ReDim Arr(1)    '2通りの可能性がある
            Arr(0) = myYear & "03"
            Arr(1) = myYear & "06"
        Case 4, 5
            ReDim Arr(0)
            Arr(0) = myYear & "06"
        Case 6
            ReDim Arr(1)    '2通りの可能性がある
            Arr(0) = myYear & "06"
            Arr(1) = myYear & "09"
        Case 7, 8
            ReDim Arr(0)
            Arr(0) = myYear & "09"
        Case 9
            ReDim Arr(1)    '2通りの可能性がある
            Arr(0) = myYear & "09"
            Arr(1) = myYear & "12"
        Case 10, 11
            ReDim Arr(0)
            Arr(0) = myYear & "12"
        Case 12
            ReDim Arr(1)    '2通りの可能性がある
            Arr(0) = myYear & "12"
            Arr(1) = myYear + 1 & "03" '12月は【翌年】の3月限もある
    End Select
    
    ArrGengetsu = Arr
End Property

Private Sub iASYNCGATE_GATE(Args_ As Variant)
    'ディレイ付き非同期処理
    With mySheet
        Select Case ArrGengetsuLength
            Case 1
                'フォームへ表示
                FormSub.addListItem "候補1件" & ArrGengetsu(0)
                
                '候補が１つだけなのでそれを採用
                Me.BrandCode = .Cells(1, 1)
                Me.Expantion = .Cells(1, 2)
                Me.BrandName = .Cells(1, 3)
                
            Case 2
                'フォームへ表示
                FormSub.addListItem "候補2件" & ArrGengetsu(0) & ":" & ArrGengetsu(1)
                FormSub.addListItem "期近残存=" & .Cells(1, 2) & ",期先残存=" & .Cells(2, 2)
                
                '期近の残存日数の欄Cells(1, 2)が空白でないときは期近を採用
                If Val(.Cells(1, 2)) > 0 Then
                    Me.BrandCode = .Cells(1, 1)
                    Me.Expantion = .Cells(1, 2)
                    Me.BrandName = .Cells(1, 3)
'                    Debug.Print "期近"
                    'フォームへ表示
                    FormSub.addListItem "期近を採用 " & BrandName
                Else
                    '期先を採用
                    Me.BrandCode = .Cells(2, 1)
                    Me.Expantion = .Cells(2, 2)
                    Me.BrandName = .Cells(2, 3)
'                    Debug.Print "期先"
                    'フォームへ表示
                    FormSub.addListItem "期先を採用 " & BrandName
                End If
                
        End Select
        
        
        '最終チェック
        '選択した先物コードに問題がないか
        '残存日数が0あるいはEmptyなら正常ではない
        If Val(Me.Expantion) = 0 Then
            '異常取得と判定
            Debug.Print myFutureCode.PetName & ":" & Me.BrandCode & ":" & Me.Expantion & " End"
            Call FormSub.addListItem("コード解決エラー:" & myFutureCode.PetName & ":" & Me.BrandCode & ":" & Me.Expantion)

        Else
            '正常取得と判断
            'フォームへ
            Call FormSub.addListItem(myFutureCode.PetName & ":" & Me.BrandCode & ":" & Me.Expantion)
        
            '--------------
            '元のFutureCodeインスタンスへコールバック
            Call myFutureCode.Resolved(Me)
        End If
        'このワークシートを削除する
        Application.DisplayAlerts = False 'アラート不要
        mySheet.Delete
        
    End With
End Sub

