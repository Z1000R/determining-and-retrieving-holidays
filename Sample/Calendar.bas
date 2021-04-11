Attribute VB_Name = "Calendar"
Option Explicit

'//////////////////////////////////////////////////
'                   概　　要
'//////////////////////////////////////////////////
'
'用　　　途：カレンダー作成
'備　　　考：休日判定処理クラス（CCompanyHoliday）が別途必要です。
'処理対象日：休日判定処理クラス（CCompanyHoliday）に依存します。

'//////////////////////////////////////////////////
'                   定数
'//////////////////////////////////////////////////

'--------------------------------------------------
'               カレンダー共通
'--------------------------------------------------

'カレンダーを作成するワークシートのプレフィックス
Private Const TARGET_SHEET_PREFIX   As String = "Calendar"

'休日色
Private Const HOLIDDAY_BACK_COLOR   As Long = vbRed
Private Const HOLIDDAY_FORE_COLOR   As Long = vbWhite


'--------------------------------------------------
'               カレンダー（BOX）
'--------------------------------------------------

'カレンダー書き込み基準セル（左上）
Private Const REFERENCE_ROW         As Long = 2
Private Const REFERENCE_COL         As Long = 2

'ヘッダ行数（月、曜日）
Private Const HEADER_ROWS           As Long = 3

'月間隔（行）
Private Const LINE_SPACING_MONTH    As Long = 2
'月間隔（列）
Private Const COLUMN_SPACING_MONTH  As Long = 2

'日間隔（行）
Private Const LINE_SPACING_DAY      As Long = 0
'日間隔（列）
Private Const COLUMN_SPACING_DAY    As Long = 0

'１行に表示する月数
Private Const MONTHS_IN_LINE        As Long = 4

'カレンダー書き込みモード
Private Enum CalendarPrintMode
    enm01To12
    enm04To03
End Enum

'カレンダーセルのサイズ
Private Const CALENDAR_COL_WIDTH    As Double = 3.25
Private Const CALENDER_ROW_HEIGHT   As Double = 14.25


'--------------------------------------------------
'               カレンダー（縦）
'--------------------------------------------------

'カレンダー書き込み基準セル（左上）
Private Const REFERENCE_ROW_V       As Long = 2
Private Const REFERENCE_COL_V       As Long = 1

'項目の表示列（基準セルからの列方向オフセット）
Private Const CALENDER_V_DATE_COL_INDEX           As Long = 0
Private Const CALENDER_V_WEEKDAY_COL_INDEX        As Long = 1
Private Const CALENDER_V_HOLIDAY_NAME_COL_INDEX   As Long = 2

'--------------------------------------------------
'               カレンダー（横）
'--------------------------------------------------

'カレンダー書き込み基準セル（左上）
Private Const REFERENCE_ROW_H       As Long = 2
Private Const REFERENCE_COL_H       As Long = 2

'項目の表示列（基準セルからの行方向オフセット）
Private Const CALENDER_H_MONTH_COL_INDEX    As Long = 0
Private Const CALENDER_H_DAY_COL_INDEX      As Long = 1
Private Const CALENDER_H_WEEKDAY_COL_INDEX  As Long = 2

Public Sub createCalendar()

    Call createCalendarY(Year(Now))
'    Call createCalendarYD(Year(Now))
'    Call createCalendarYMV(2021, 4, 2, False)
'    Call createCalendarYMH(2021, 4, 2)

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'               パブリックメソッド
'++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'
'機　　　能：指定年の１月から１２月までのカレンダー作成
'パラメータ：
'           lYear：作成するカレンダーの年
'備　　　考：
'
'//////////////////////////////////////////////////
Public Sub createCalendarY(ByVal lYear As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    'カレンダーを作成するワークシート取得
    Set ws = getTargetSheet(lYear)

    Set cch = New CCompanyHoliday

    For i = 1 To 12
        '月の基準位置を取得
        Set r = getReferenceRange(ws, i, enm01To12)

        '指定月のカレンダー書き込み
        Call createCalendarSub(cch, r, lYear, i)
    Next i

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'機　　　能：指定年度の４月から３月までのカレンダー作成
'パラメータ：
'           lYear：作成するカレンダーの年度
'備　　　考：
'
'//////////////////////////////////////////////////
Public Sub createCalendarYD(ByVal lYear As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    'カレンダーを作成するワークシート取得
    Set ws = getTargetSheet(lYear)

    Set cch = New CCompanyHoliday

    For i = 4 To 12
        '月の基準位置を取得
        Set r = getReferenceRange(ws, i, enm04To03)

        '指定月のカレンダー書き込み
        Call createCalendarSub(cch, r, lYear, i)
    Next i

    For i = 1 To 3
        '月の基準位置を取得
        Set r = getReferenceRange(ws, i, enm04To03)

        '指定月のカレンダー書き込み
        Call createCalendarSub(cch, r, lYear + 1, i)
    Next i

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'機　　　能：指定年月の縦カレンダー作成
'パラメータ：
'           lYear           ：作成するカレンダーの最初の月の年
'           lBeginMonth     ：作成するカレンダーの最初の月
'           lMonthes        ：作成するカレンダーの月数
'           printHolidayName：休日名を出力するか
'                               True ：出力する
'                               False：出力しない
'備　　　考：
'
'//////////////////////////////////////////////////
Public Sub createCalendarYMV(ByVal lYear As Long, _
                             ByVal lBeginMonth As Long, _
                             ByVal lMonthes As Long, _
                             Optional ByVal printHolidayName As Boolean = False)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    'カレンダーを作成するワークシート取得
    Set ws = getTargetSheet(lYear)

    Set r = ws.Cells(REFERENCE_ROW_V, REFERENCE_COL_V)

    Set cch = New CCompanyHoliday

    '指定月のカレンダー書き込み
    Call createCalendarVSub(cch, r, lYear, lBeginMonth, lMonthes, printHolidayName)

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'機　　　能：指定年月の横カレンダー作成
'パラメータ：
'           lYear           ：作成するカレンダーの最初の月の年
'           lBeginMonth     ：作成するカレンダーの最初の月
'           lMonthes        ：作成するカレンダーの月数
'備　　　考：
'
'//////////////////////////////////////////////////
Public Sub createCalendarYMH(ByVal lYear As Long, _
                             ByVal lBeginMonth As Long, _
                             ByVal lMonthes As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    'カレンダーを作成するワークシート取得
    Set ws = getTargetSheet(lYear)

    Set r = ws.Cells(REFERENCE_ROW_H, REFERENCE_COL_H)

    Set cch = New CCompanyHoliday

    '指定月のカレンダー書き込み
    Call createCalendarHSub(cch, r, lYear, lBeginMonth, lMonthes)

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'               プライベートメソッド
'++++++++++++++++++++++++++++++++++++++++++++++++++

'カレンダーを作成するワークシートの取得
Private Function getTargetSheet(ByVal lYear As Long) As Worksheet

    Dim ws          As Worksheet
    Dim s           As Worksheet
    Dim sSheetName  As String

    'カレンダーを作成するシート名
    sSheetName = TARGET_SHEET_PREFIX & CStr(lYear)

    '指定シート名を検索
    For Each s In ThisWorkbook.Worksheets
        If s.Name = sSheetName Then
            '見つかったら変数にセットし、クリア
            Set ws = ThisWorkbook.Worksheets(sSheetName)

            With ws
                .Range(.Cells(1, 1), .Cells(.Rows.Count, .Columns.Count)).Clear
            End With

            Exit For
        End If
    Next s

    If ws Is Nothing Then
        '見つからなかったらシートを作成
        With ThisWorkbook.Worksheets
            Set ws = .Add(after:=ThisWorkbook.Worksheets(.Count))
        End With

        'リネーム
        ws.Name = sSheetName
    End If

    With ws.Cells
        .ColumnWidth = CALENDAR_COL_WIDTH
        .RowHeight = CALENDER_ROW_HEIGHT
    End With

    Set getTargetSheet = ws

End Function

'カレンダー書き込み
Private Sub createCalendarSub(ByRef cch As CCompanyHoliday, _
                              ByRef r As Range, _
                              ByVal lYear As Long, _
                              ByVal lMonth As Long)

    Dim dtBegin As Date
    Dim dtEnd   As Date
    Dim dtDate  As Date
    Dim lDays   As Long
    Dim lRowIndex   As Long
    Dim lColIndex   As Long
    Dim lRowOffset  As Long
    Dim lColOffset  As Long
    Dim i       As Long

    Const WEEKDAYS  As String = "日月火水木金土"

    dtBegin = DateSerial(lYear, lMonth, 1)
    dtEnd = DateSerial(lYear, lMonth + 1, 0)

    '月
    r.Value = StrConv(CStr(lMonth), vbWide) & "月"

    '曜日
    lRowOffset = HEADER_ROWS - 1

    For i = 0 To 6
        lColOffset = i * (COLUMN_SPACING_DAY + 1)
        r.Offset(lRowOffset, lColOffset).Value = Mid$(WEEKDAYS, i + 1, 1)
    Next i

    lRowIndex = 0
    lColIndex = Weekday(dtBegin) - vbSunday

    'カレンダー書き込み
    With r
        For i = 0 To Day(dtEnd) - 1
            '基準位置からのオフセット計算
            lRowOffset = lRowIndex * (LINE_SPACING_DAY + 1) + HEADER_ROWS
            lColOffset = lColIndex * (COLUMN_SPACING_DAY + 1)

            '日付書き込み
            .Offset(lRowOffset, lColOffset).Value = i + 1

            dtDate = DateAdd("d", i, dtBegin)

            '休日判定
            If cch.isCompanyHoliday(dtDate) Then
                .Offset(lRowOffset, lColOffset).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(lRowOffset, lColOffset).Font.Color = HOLIDDAY_FORE_COLOR
            End If

            lColIndex = lColIndex + 1

            If lColIndex Mod 7 = 0 Then
                lColIndex = 0

                lRowIndex = lRowIndex + 1
            End If
        Next i
    End With

End Sub

'カレンダー（縦）書き込み
Private Sub createCalendarVSub(ByRef cch As CCompanyHoliday, _
                               ByRef r As Range, _
                               ByVal lYear As Long, _
                               ByVal lBeginMonth As Long, _
                               ByVal lMonthes As Long, _
                               ByVal printHolidayName As Boolean)

    Dim dtBegin As Date
    Dim dtEnd   As Date
    Dim dtDate  As Date
    Dim lDays   As Long
    Dim sHolidayName    As String
    Dim i       As Long

    Const WEEKDAYS  As String = "日月火水木金土"

    dtBegin = DateSerial(lYear, lBeginMonth, 1)
    dtEnd = DateSerial(lYear, lBeginMonth + lMonthes, 0)

    lDays = DateDiff("d", dtBegin, dtEnd) + 1

    'カレンダー書き込み
    With r
        For i = 0 To lDays - 1
            dtDate = DateAdd("d", i, dtBegin)

            '日付書き込み
            .Offset(i, CALENDER_V_DATE_COL_INDEX).Value = dtDate
            .Offset(i, CALENDER_V_WEEKDAY_COL_INDEX).Value = Mid$(WEEKDAYS, Weekday(dtDate), 1)

            '休日判定
            If cch.isCompanyHoliday2(dtDate, sHolidayName) Then
                '日付（色をつけたくない場合には、以下の２行をコメントする）
                .Offset(i, CALENDER_V_DATE_COL_INDEX).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(i, CALENDER_V_DATE_COL_INDEX).Font.Color = HOLIDDAY_FORE_COLOR

                '曜日（色をつけたくない場合には、以下の２行をコメントする）
                .Offset(i, CALENDER_V_WEEKDAY_COL_INDEX).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(i, CALENDER_V_WEEKDAY_COL_INDEX).Font.Color = HOLIDDAY_FORE_COLOR

                If printHolidayName Then
                    .Offset(i, CALENDER_V_HOLIDAY_NAME_COL_INDEX).Value = sHolidayName
                End If
            End If
        Next i

        .Offset(0, CALENDER_V_DATE_COL_INDEX).EntireColumn.AutoFit
        .Offset(0, CALENDER_V_WEEKDAY_COL_INDEX).EntireColumn.AutoFit
        If printHolidayName Then
            .Offset(0, CALENDER_V_HOLIDAY_NAME_COL_INDEX).EntireColumn.AutoFit
        End If
    End With

End Sub

'カレンダー（横）書き込み
Private Sub createCalendarHSub(ByRef cch As CCompanyHoliday, _
                               ByRef r As Range, _
                               ByVal lYear As Long, _
                               ByVal lBeginMonth As Long, _
                               ByVal lMonthes As Long)

    Dim dtBegin As Date
    Dim dtEnd   As Date
    Dim dtDate  As Date
    Dim lDays   As Long
    Dim sHolidayName    As String
    Dim rDate   As Range
    Dim i       As Long

    Const WEEKDAYS  As String = "日月火水木金土"

    dtBegin = DateSerial(lYear, lBeginMonth, 1)
    dtEnd = DateSerial(lYear, lBeginMonth + lMonthes, 0)

    lDays = DateDiff("d", dtBegin, dtEnd) + 1

    'カレンダー書き込み
    With r
        For i = 0 To lDays - 1
            dtDate = DateAdd("d", i, dtBegin)

            '日付書き込み
            .Offset(CALENDER_H_MONTH_COL_INDEX, i).Value = dtDate
            .Offset(CALENDER_H_DAY_COL_INDEX, i).Value = dtDate
            .Offset(CALENDER_H_WEEKDAY_COL_INDEX, i).Value = Mid$(WEEKDAYS, Weekday(dtDate), 1)

            '休日判定
            If cch.isCompanyHoliday2(dtDate, sHolidayName) Then
                '日（色をつけたくない場合には、以下の２行をコメントする）
                .Offset(CALENDER_H_DAY_COL_INDEX, i).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(CALENDER_H_DAY_COL_INDEX, i).Font.Color = HOLIDDAY_FORE_COLOR

                '曜日（色をつけたくない場合には、以下の２行をコメントする）
                .Offset(CALENDER_H_WEEKDAY_COL_INDEX, i).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(CALENDER_H_WEEKDAY_COL_INDEX, i).Font.Color = HOLIDDAY_FORE_COLOR
            End If
        Next i
    End With

    Set rDate = r.Resize(1, lDays)

    With rDate.Offset(CALENDER_H_MONTH_COL_INDEX)
        .HorizontalAlignment = xlCenter
        .NumberFormatLocal = "m"
    End With

    With rDate.Offset(CALENDER_H_DAY_COL_INDEX)
        .HorizontalAlignment = xlCenter
        .NumberFormatLocal = "d"
    End With

    With rDate.Offset(CALENDER_H_WEEKDAY_COL_INDEX)
        .HorizontalAlignment = xlCenter
    End With

End Sub

'指定年月のカレンダー書き込み基準位置取得
Private Function getReferenceRange(ByRef ws As Worksheet, _
                                   ByVal lMonth As Long, _
                                   ByVal lPrintMode As CalendarPrintMode) As Range

    '１ヶ月の最大週数
    Const MAX_WEEKS_IN_MONTH    As Long = 6

    Dim lLinesInMonth   As Long
    Dim lColsInMonth    As Long
    Dim lMonthW As Long
    Dim lNthV   As Long
    Dim lNthH   As Long
    Dim lRow    As Long
    Dim lCol    As Long

    '１ヶ月を表示するのに必要な行数
    lLinesInMonth = HEADER_ROWS + MAX_WEEKS_IN_MONTH + (MAX_WEEKS_IN_MONTH - 1) * LINE_SPACING_DAY

    If lPrintMode = CalendarPrintMode.enm04To03 Then
        Select Case lMonth
        Case 1 To 3
            lMonthW = lMonth + 9
        Case 4 To 12
            lMonthW = lMonth - 3
        End Select
    Else
        lMonthW = lMonth
    End If

    '指定月が何段目に表示されるか
    lNthV = (lMonthW + MONTHS_IN_LINE - 1) \ MONTHS_IN_LINE

    lRow = REFERENCE_ROW + (lNthV - 1) * (lLinesInMonth + LINE_SPACING_MONTH)

    '１ヶ月を表示するのに必要な列数
    lColsInMonth = 7 + (7 - 1) * COLUMN_SPACING_DAY

    '指定月が何列目に表示されるか
    lNthH = lMonthW Mod MONTHS_IN_LINE

    If lNthH Mod MONTHS_IN_LINE = 0 Then
        lNthH = MONTHS_IN_LINE
    End If

    lCol = REFERENCE_COL + (lNthH - 1) * (lColsInMonth + COLUMN_SPACING_MONTH)

    Set getReferenceRange = ws.Cells(lRow, lCol)

End Function


