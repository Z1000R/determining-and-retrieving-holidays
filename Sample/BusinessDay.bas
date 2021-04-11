Attribute VB_Name = "BusinessDay"
Option Explicit


'//////////////////////////////////////////////////
'                   概　　要
'//////////////////////////////////////////////////
'
'用　　　途：営業日取得
'備　　　考：休日判定処理クラス（CCompanyHoliday）が別途必要です。
'処理対象日：休日判定処理クラス（CCompanyHoliday）に依存します。

'//////////////////////////////////////////////////
'                   変数
'//////////////////////////////////////////////////

Private cch_    As CCompanyHoliday


'++++++++++++++++++++++++++++++++++++++++++++++++++
'               パブリックメソッド
'++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'
'機　　　能：指定日を基準に、指定営業日数だけ移動した日付を取得する
'パラメータ：
'           dtBegin ：基準日
'           lDays   ：移動する日数（負数でも可）
'           dtResult：移動した日付
'復　帰　値：
'備　　　考：
'
'//////////////////////////////////////////////////
Public Function getNthWorkingDay(ByVal dtBegin As Date, ByVal lDays As Long, ByRef dtResult As Date) As Boolean

    Const VALID_FIRST_YEAR  As Long = 1948

    Dim dtBeginW    As Date
    Dim dtTemp      As Date
    Dim lAddedDays  As Long
    Dim lWorkingDays    As Long
    Dim lStep       As Long
    Dim lInitializedYear    As Long

    getNthWorkingDay = True

    dtBeginW = DateSerial(Year(dtBegin), Month(dtBegin), Day(dtBegin))

    If lDays = 0 Then
        dtResult = dtBeginW

        Exit Function
    End If

    If cch_ Is Nothing Then
        Set cch_ = New CCompanyHoliday
    End If

    lInitializedYear = cch_.InitializedLastYear

    lAddedDays = 0

    lStep = Sgn(lDays)

    Do Until lWorkingDays = lDays

        lAddedDays = lAddedDays + lStep

        dtTemp = DateAdd("d", lAddedDays, dtBeginW)

        If Year(dtTemp) > lInitializedYear Then
            lInitializedYear = Year(dtTemp)

            Call cch_.reInitialize(lInitializedYear)
        ElseIf Year(dtTemp) <= VALID_FIRST_YEAR Then
            '「国民の祝日に関する法律」施行年以前ならエラーとする
            '厳密には1948/7/20施行であるが、
            '簡略化のため1948/12/31以前ならエラーにしている
            getNthWorkingDay = False

            Exit Function
        End If

        If cch_.isCompanyHoliday(dtTemp) = False Then
            lWorkingDays = lWorkingDays + lStep
        End If
    Loop

    dtResult = DateAdd("d", lAddedDays, dtBeginW)

    Set cch_ = Nothing

End Function

Public Sub 第N営業日取得()

    Dim d1 As Date
    Dim d2 As Date
    Dim diff As Long

    d1 = #4/28/2021#

    diff = 5
    Call getNthWorkingDay(d1, diff, d2)
    Debug.Print d1, diff, d2

    diff = -1
    Call getNthWorkingDay(d1, diff, d2)
    Debug.Print d1, diff, d2

End Sub
