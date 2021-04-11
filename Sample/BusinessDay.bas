Attribute VB_Name = "BusinessDay"
Option Explicit


'//////////////////////////////////////////////////
'                   �T�@�@�v
'//////////////////////////////////////////////////
'
'�p�@�@�@�r�F�c�Ɠ��擾
'���@�@�@�l�F�x�����菈���N���X�iCCompanyHoliday�j���ʓr�K�v�ł��B
'�����Ώۓ��F�x�����菈���N���X�iCCompanyHoliday�j�Ɉˑ����܂��B

'//////////////////////////////////////////////////
'                   �ϐ�
'//////////////////////////////////////////////////

Private cch_    As CCompanyHoliday


'++++++++++++++++++++++++++++++++++++++++++++++++++
'               �p�u���b�N���\�b�h
'++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'
'�@�@�@�@�\�F�w�������ɁA�w��c�Ɠ��������ړ��������t���擾����
'�p�����[�^�F
'           dtBegin �F���
'           lDays   �F�ړ���������i�����ł��j
'           dtResult�F�ړ��������t
'���@�A�@�l�F
'���@�@�@�l�F
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
            '�u�����̏j���Ɋւ���@���v�{�s�N�ȑO�Ȃ�G���[�Ƃ���
            '�����ɂ�1948/7/20�{�s�ł��邪�A
            '�ȗ����̂���1948/12/31�ȑO�Ȃ�G���[�ɂ��Ă���
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

Public Sub ��N�c�Ɠ��擾()

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
