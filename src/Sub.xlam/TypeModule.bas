Attribute VB_Name = "TypeModule"
Option Explicit

'構造体 (構造体はプロシージャ内では定義できない)
'Typeステートメント:ユーザー定義のデータ型を作成する
' 要素の値はLongとして評価される
'***************************
'列挙型（列挙体）Enum
'***************************
'8/28追記


Public Enum En_Buysell
    売 = 1
    買 = 3
End Enum

'注文区分 : {0:通常注文 , 1:逆指値付通常注文 , 2:逆指値注文}  (Modifyで訂正元の注文区分を訂正する場合､{3: 通常注文⇒逆指値付通常注文､4: 逆指値付通常注文⇒通常注文}
Public Enum En_Ordertype
    通常注文 = 0
    逆指値付通常注文 = 1
    逆指値注文 = 2
    通常注文⇒逆指値付通常注文 = 3
    逆指値付通常注文⇒通常注文 = 4
End Enum
Public Enum En_PriceType
    成行 = 0
    指値 = 1
End Enum
Public Enum En_ExecQuantityType
    FOK = 1     '一部約定後に未執行数量が残るとき、その残数量を有効とする条件
    'FAK = 2    '(<==中途半端。使用しない) 一部約定後に未執行数量が残るとき、その残数量を失効させる条件
    FAS = 3     '全数量が直ちに約定しない場合は、その全数量を失効させる条件
End Enum
Public Enum En_ExpiryType
    当セッション = 1
    引け = 4
    期間限定 = 5
    最終取引日 = 9
End Enum
Public Enum En_StopTriggerType
    以上 = 1
    以下 = 2
End Enum
Public Enum En_StopPriceType
    成行 = 0
    指値 = 1
End Enum

'V11
Public Enum En_TrailType
    なし = 0
    トレイリングストップ = 1
    ロスカット値のみ = 2
End Enum

Public Enum En_ModifyType
    価格のみ = 1
    逆指値のみ = 2
    価格および逆指値 = 3
End Enum

'2024/10/14追記
Public Enum En_CloseType
    なし = 0
    自動クローズ = 1
    成行EXIT = 2
End Enum

Public Enum En_TargetDataType
    データなし = 0
    初期更新 = 1
    新規 = 2
    更新 = 3
    削除 = 4
End Enum
'2025/12/9追加
Public Enum En_DealType
    DummyOrder = 0
    FutureOrder = 1
    StockOrder = 2
End Enum
'***************************
'ユーザー定義型・構造体（Type） :留意点：イベントの引数にはTypeオブジェクトは使用できない
'***************************
'なし 2025/3/9現在

