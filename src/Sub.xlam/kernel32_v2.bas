Attribute VB_Name = "kernel32_v2"
Option Explicit

'GetLocalTimeで使用する構造体
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'マイクロ秒の計測に使用
Private Type LARGE_INTEGER
    QuadPart As LongLong
End Type

'Declareステートメント:ダイナミックリンクライブラリ（DLL）の外部プロシージャへの参照を宣言します。
'funcion , sub よりも上部に記述する必要があります
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare PtrSafe Sub getLocalTime Lib "kernel32" Alias "GetLocalTime" (lpSystemTime As SYSTEMTIME)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

'2024/9/26追記
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal LENGTH As LongPtr)

'*************************************
'**ミリ秒単位で時間を取得する**
'2023/3/10
'kamogashira
'GetLocalTime関数はkernel32ライブラリで定義されている外部プロシージャのため､利用する場合にはDeclareステートメントで利用を宣言しておく必要があります｡
'また､GetLocalTime関数は日時をSYSTEMTIME構造体で返すため､SYSTEMTIME構造体も一緒に定義しておきます｡
'*************************************

'【日時】
'ミリ秒(1/1000秒)単位の現在日時の文字列を返す
'YYYY/MM/DD HH:NN:SS FFF (FFFはミリ秒)
Public Function getDateTimeStamp() As String
    Dim SysTime As SYSTEMTIME
    Dim myStr As String
    '現在日時取得
    Call getLocalTime(SysTime)
    myStr = Format(SysTime.wYear, "0000")
    myStr = myStr & "/"
    myStr = myStr & Format(SysTime.wMonth, "00")
    myStr = myStr & "/"
    myStr = myStr & Format(SysTime.wDay, "00")
    myStr = myStr & " "
    myStr = myStr & Format(SysTime.wHour, "00")
    myStr = myStr & ":"
    myStr = myStr & Format(SysTime.wMinute, "00")
    myStr = myStr & ":"
    myStr = myStr & Format(SysTime.wSecond, "00")
    myStr = myStr & " "
    myStr = myStr & Format(SysTime.wMilliseconds, "000")
     
    getDateTimeStamp = myStr
End Function

'【時刻】
'ミリ秒(1/1000秒)単位の現在[時刻]の文字列を返す
'HH:NN:SS FFF (FFFはミリ秒)
Public Function getTimeStamp() As String
    Dim SysTime As SYSTEMTIME
    Dim myStr As String
    '現在日時取得
    Call getLocalTime(SysTime)
    myStr = Format(SysTime.wHour, "00")
    myStr = myStr & ":"
    myStr = myStr & Format(SysTime.wMinute, "00")
    myStr = myStr & ":"
    myStr = myStr & Format(SysTime.wSecond, "00")
    myStr = myStr & " "
    myStr = myStr & Format(SysTime.wMilliseconds, "000")
     
    getTimeStamp = myStr
End Function

'マイクロ秒単位で経過時間を取得
'参考：https://vbabeginner.net/measure-milliseconds-and-microseconds/
'マイクロ秒は1秒の100万分の1(1/1,000,000秒)
Public Function getMicroSecond() As Double
    Dim procTime            As LARGE_INTEGER '// 高分解能パフォーマンスカウンタ値（システム起動からの加算値）
    Dim frequency           As LARGE_INTEGER '// 高分解能パフォーマンスカウンタの周波数（１秒間に増えるカウントの数）
    
    '// 計測時刻を0で初期化
    getMicroSecond = 0

    '// 更新頻度を取得
    Call QueryPerformanceFrequency(frequency)

    '// 処理時刻を取得
    Call QueryPerformanceCounter(procTime)

    '// カウンタ値を１秒間のカウント増加数で割り、正確な時刻を算出
    '// GetMicroSecond = procTime / frequency '// 32bit版
    getMicroSecond = procTime.QuadPart / frequency.QuadPart
End Function
