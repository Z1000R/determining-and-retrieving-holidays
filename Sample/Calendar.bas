Attribute VB_Name = "Calendar"
Option Explicit

'//////////////////////////////////////////////////
'                   �T�@�@�v
'//////////////////////////////////////////////////
'
'�p�@�@�@�r�F�J�����_�[�쐬
'���@�@�@�l�F�x�����菈���N���X�iCCompanyHoliday�j���ʓr�K�v�ł��B
'�����Ώۓ��F�x�����菈���N���X�iCCompanyHoliday�j�Ɉˑ����܂��B

'//////////////////////////////////////////////////
'                   �萔
'//////////////////////////////////////////////////

'--------------------------------------------------
'               �J�����_�[����
'--------------------------------------------------

'�J�����_�[���쐬���郏�[�N�V�[�g�̃v���t�B�b�N�X
Private Const TARGET_SHEET_PREFIX   As String = "Calendar"

'�x���F
Private Const HOLIDDAY_BACK_COLOR   As Long = vbRed
Private Const HOLIDDAY_FORE_COLOR   As Long = vbWhite


'--------------------------------------------------
'               �J�����_�[�iBOX�j
'--------------------------------------------------

'�J�����_�[�������݊�Z���i����j
Private Const REFERENCE_ROW         As Long = 2
Private Const REFERENCE_COL         As Long = 2

'�w�b�_�s���i���A�j���j
Private Const HEADER_ROWS           As Long = 3

'���Ԋu�i�s�j
Private Const LINE_SPACING_MONTH    As Long = 2
'���Ԋu�i��j
Private Const COLUMN_SPACING_MONTH  As Long = 2

'���Ԋu�i�s�j
Private Const LINE_SPACING_DAY      As Long = 0
'���Ԋu�i��j
Private Const COLUMN_SPACING_DAY    As Long = 0

'�P�s�ɕ\�����錎��
Private Const MONTHS_IN_LINE        As Long = 4

'�J�����_�[�������݃��[�h
Private Enum CalendarPrintMode
    enm01To12
    enm04To03
End Enum

'�J�����_�[�Z���̃T�C�Y
Private Const CALENDAR_COL_WIDTH    As Double = 3.25
Private Const CALENDER_ROW_HEIGHT   As Double = 14.25


'--------------------------------------------------
'               �J�����_�[�i�c�j
'--------------------------------------------------

'�J�����_�[�������݊�Z���i����j
Private Const REFERENCE_ROW_V       As Long = 2
Private Const REFERENCE_COL_V       As Long = 1

'���ڂ̕\����i��Z������̗�����I�t�Z�b�g�j
Private Const CALENDER_V_DATE_COL_INDEX           As Long = 0
Private Const CALENDER_V_WEEKDAY_COL_INDEX        As Long = 1
Private Const CALENDER_V_HOLIDAY_NAME_COL_INDEX   As Long = 2

'--------------------------------------------------
'               �J�����_�[�i���j
'--------------------------------------------------

'�J�����_�[�������݊�Z���i����j
Private Const REFERENCE_ROW_H       As Long = 2
Private Const REFERENCE_COL_H       As Long = 2

'���ڂ̕\����i��Z������̍s�����I�t�Z�b�g�j
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
'               �p�u���b�N���\�b�h
'++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'
'�@�@�@�@�\�F�w��N�̂P������P�Q���܂ł̃J�����_�[�쐬
'�p�����[�^�F
'           lYear�F�쐬����J�����_�[�̔N
'���@�@�@�l�F
'
'//////////////////////////////////////////////////
Public Sub createCalendarY(ByVal lYear As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    '�J�����_�[���쐬���郏�[�N�V�[�g�擾
    Set ws = getTargetSheet(lYear)

    Set cch = New CCompanyHoliday

    For i = 1 To 12
        '���̊�ʒu���擾
        Set r = getReferenceRange(ws, i, enm01To12)

        '�w�茎�̃J�����_�[��������
        Call createCalendarSub(cch, r, lYear, i)
    Next i

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'�@�@�@�@�\�F�w��N�x�̂S������R���܂ł̃J�����_�[�쐬
'�p�����[�^�F
'           lYear�F�쐬����J�����_�[�̔N�x
'���@�@�@�l�F
'
'//////////////////////////////////////////////////
Public Sub createCalendarYD(ByVal lYear As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    '�J�����_�[���쐬���郏�[�N�V�[�g�擾
    Set ws = getTargetSheet(lYear)

    Set cch = New CCompanyHoliday

    For i = 4 To 12
        '���̊�ʒu���擾
        Set r = getReferenceRange(ws, i, enm04To03)

        '�w�茎�̃J�����_�[��������
        Call createCalendarSub(cch, r, lYear, i)
    Next i

    For i = 1 To 3
        '���̊�ʒu���擾
        Set r = getReferenceRange(ws, i, enm04To03)

        '�w�茎�̃J�����_�[��������
        Call createCalendarSub(cch, r, lYear + 1, i)
    Next i

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'�@�@�@�@�\�F�w��N���̏c�J�����_�[�쐬
'�p�����[�^�F
'           lYear           �F�쐬����J�����_�[�̍ŏ��̌��̔N
'           lBeginMonth     �F�쐬����J�����_�[�̍ŏ��̌�
'           lMonthes        �F�쐬����J�����_�[�̌���
'           printHolidayName�F�x�������o�͂��邩
'                               True �F�o�͂���
'                               False�F�o�͂��Ȃ�
'���@�@�@�l�F
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

    '�J�����_�[���쐬���郏�[�N�V�[�g�擾
    Set ws = getTargetSheet(lYear)

    Set r = ws.Cells(REFERENCE_ROW_V, REFERENCE_COL_V)

    Set cch = New CCompanyHoliday

    '�w�茎�̃J�����_�[��������
    Call createCalendarVSub(cch, r, lYear, lBeginMonth, lMonthes, printHolidayName)

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'//////////////////////////////////////////////////
'
'�@�@�@�@�\�F�w��N���̉��J�����_�[�쐬
'�p�����[�^�F
'           lYear           �F�쐬����J�����_�[�̍ŏ��̌��̔N
'           lBeginMonth     �F�쐬����J�����_�[�̍ŏ��̌�
'           lMonthes        �F�쐬����J�����_�[�̌���
'���@�@�@�l�F
'
'//////////////////////////////////////////////////
Public Sub createCalendarYMH(ByVal lYear As Long, _
                             ByVal lBeginMonth As Long, _
                             ByVal lMonthes As Long)

    Dim ws      As Worksheet
    Dim r       As Range
    Dim cch     As CCompanyHoliday
    Dim i       As Long

    '�J�����_�[���쐬���郏�[�N�V�[�g�擾
    Set ws = getTargetSheet(lYear)

    Set r = ws.Cells(REFERENCE_ROW_H, REFERENCE_COL_H)

    Set cch = New CCompanyHoliday

    '�w�茎�̃J�����_�[��������
    Call createCalendarHSub(cch, r, lYear, lBeginMonth, lMonthes)

    Set cch = Nothing

    Debug.Print "Done."

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++
'               �v���C�x�[�g���\�b�h
'++++++++++++++++++++++++++++++++++++++++++++++++++

'�J�����_�[���쐬���郏�[�N�V�[�g�̎擾
Private Function getTargetSheet(ByVal lYear As Long) As Worksheet

    Dim ws          As Worksheet
    Dim s           As Worksheet
    Dim sSheetName  As String

    '�J�����_�[���쐬����V�[�g��
    sSheetName = TARGET_SHEET_PREFIX & CStr(lYear)

    '�w��V�[�g��������
    For Each s In ThisWorkbook.Worksheets
        If s.Name = sSheetName Then
            '����������ϐ��ɃZ�b�g���A�N���A
            Set ws = ThisWorkbook.Worksheets(sSheetName)

            With ws
                .Range(.Cells(1, 1), .Cells(.Rows.Count, .Columns.Count)).Clear
            End With

            Exit For
        End If
    Next s

    If ws Is Nothing Then
        '������Ȃ�������V�[�g���쐬
        With ThisWorkbook.Worksheets
            Set ws = .Add(after:=ThisWorkbook.Worksheets(.Count))
        End With

        '���l�[��
        ws.Name = sSheetName
    End If

    With ws.Cells
        .ColumnWidth = CALENDAR_COL_WIDTH
        .RowHeight = CALENDER_ROW_HEIGHT
    End With

    Set getTargetSheet = ws

End Function

'�J�����_�[��������
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

    Const WEEKDAYS  As String = "�����ΐ��؋��y"

    dtBegin = DateSerial(lYear, lMonth, 1)
    dtEnd = DateSerial(lYear, lMonth + 1, 0)

    '��
    r.Value = StrConv(CStr(lMonth), vbWide) & "��"

    '�j��
    lRowOffset = HEADER_ROWS - 1

    For i = 0 To 6
        lColOffset = i * (COLUMN_SPACING_DAY + 1)
        r.Offset(lRowOffset, lColOffset).Value = Mid$(WEEKDAYS, i + 1, 1)
    Next i

    lRowIndex = 0
    lColIndex = Weekday(dtBegin) - vbSunday

    '�J�����_�[��������
    With r
        For i = 0 To Day(dtEnd) - 1
            '��ʒu����̃I�t�Z�b�g�v�Z
            lRowOffset = lRowIndex * (LINE_SPACING_DAY + 1) + HEADER_ROWS
            lColOffset = lColIndex * (COLUMN_SPACING_DAY + 1)

            '���t��������
            .Offset(lRowOffset, lColOffset).Value = i + 1

            dtDate = DateAdd("d", i, dtBegin)

            '�x������
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

'�J�����_�[�i�c�j��������
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

    Const WEEKDAYS  As String = "�����ΐ��؋��y"

    dtBegin = DateSerial(lYear, lBeginMonth, 1)
    dtEnd = DateSerial(lYear, lBeginMonth + lMonthes, 0)

    lDays = DateDiff("d", dtBegin, dtEnd) + 1

    '�J�����_�[��������
    With r
        For i = 0 To lDays - 1
            dtDate = DateAdd("d", i, dtBegin)

            '���t��������
            .Offset(i, CALENDER_V_DATE_COL_INDEX).Value = dtDate
            .Offset(i, CALENDER_V_WEEKDAY_COL_INDEX).Value = Mid$(WEEKDAYS, Weekday(dtDate), 1)

            '�x������
            If cch.isCompanyHoliday2(dtDate, sHolidayName) Then
                '���t�i�F���������Ȃ��ꍇ�ɂ́A�ȉ��̂Q�s���R�����g����j
                .Offset(i, CALENDER_V_DATE_COL_INDEX).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(i, CALENDER_V_DATE_COL_INDEX).Font.Color = HOLIDDAY_FORE_COLOR

                '�j���i�F���������Ȃ��ꍇ�ɂ́A�ȉ��̂Q�s���R�����g����j
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

'�J�����_�[�i���j��������
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

    Const WEEKDAYS  As String = "�����ΐ��؋��y"

    dtBegin = DateSerial(lYear, lBeginMonth, 1)
    dtEnd = DateSerial(lYear, lBeginMonth + lMonthes, 0)

    lDays = DateDiff("d", dtBegin, dtEnd) + 1

    '�J�����_�[��������
    With r
        For i = 0 To lDays - 1
            dtDate = DateAdd("d", i, dtBegin)

            '���t��������
            .Offset(CALENDER_H_MONTH_COL_INDEX, i).Value = dtDate
            .Offset(CALENDER_H_DAY_COL_INDEX, i).Value = dtDate
            .Offset(CALENDER_H_WEEKDAY_COL_INDEX, i).Value = Mid$(WEEKDAYS, Weekday(dtDate), 1)

            '�x������
            If cch.isCompanyHoliday2(dtDate, sHolidayName) Then
                '���i�F���������Ȃ��ꍇ�ɂ́A�ȉ��̂Q�s���R�����g����j
                .Offset(CALENDER_H_DAY_COL_INDEX, i).Interior.Color = HOLIDDAY_BACK_COLOR
                .Offset(CALENDER_H_DAY_COL_INDEX, i).Font.Color = HOLIDDAY_FORE_COLOR

                '�j���i�F���������Ȃ��ꍇ�ɂ́A�ȉ��̂Q�s���R�����g����j
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

'�w��N���̃J�����_�[�������݊�ʒu�擾
Private Function getReferenceRange(ByRef ws As Worksheet, _
                                   ByVal lMonth As Long, _
                                   ByVal lPrintMode As CalendarPrintMode) As Range

    '�P�����̍ő�T��
    Const MAX_WEEKS_IN_MONTH    As Long = 6

    Dim lLinesInMonth   As Long
    Dim lColsInMonth    As Long
    Dim lMonthW As Long
    Dim lNthV   As Long
    Dim lNthH   As Long
    Dim lRow    As Long
    Dim lCol    As Long

    '�P������\������̂ɕK�v�ȍs��
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

    '�w�茎�����i�ڂɕ\������邩
    lNthV = (lMonthW + MONTHS_IN_LINE - 1) \ MONTHS_IN_LINE

    lRow = REFERENCE_ROW + (lNthV - 1) * (lLinesInMonth + LINE_SPACING_MONTH)

    '�P������\������̂ɕK�v�ȗ�
    lColsInMonth = 7 + (7 - 1) * COLUMN_SPACING_DAY

    '�w�茎������ڂɕ\������邩
    lNthH = lMonthW Mod MONTHS_IN_LINE

    If lNthH Mod MONTHS_IN_LINE = 0 Then
        lNthH = MONTHS_IN_LINE
    End If

    lCol = REFERENCE_COL + (lNthH - 1) * (lColsInMonth + COLUMN_SPACING_MONTH)

    Set getReferenceRange = ws.Cells(lRow, lCol)

End Function


