VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCompanyHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////
'                   �T�@�@�v
'//////////////////////////////////////////////////
'
'�p�@�@�@�r�F��Ћx���擾�A�m�F�p
'�����Ώۓ��F1948/7/20�ȍ~�i2050�N�܂ł́A�t���̓��A�H���̓��m�F�ς݁j
'���@�@�@�l�F���������ܗցE�p�������s�b�N���ʑ[�u�@�i2018/6/13�Q�c�@�{��c�����j�Ή��ς�
'��@�@�@���F2018/ 9/ 9
'���@�@�@��F2019/ ?/ ?
'            �V�c�a�����i2020/2/23�j�Ή��ς�
'            2020/11/28
'            �u�����O�\��N�����I�����s�b�N���Z���E�����p�������s�b�N���Z�����ʑ[�u�@���̈ꕔ����������@���̎{�s�v�ւ̑Ή�


'//////////////////////////////////////////////////
'                   �Q�Ɛݒ�
'//////////////////////////////////////////////////

'Dictionary�p
'Microsoft Scripting Runtime


'//////////////////////////////////////////////////
'                ���[�U�[��`�^
'//////////////////////////////////////////////////

'�����Œ�̏j�����
Private Type FixMD
    sMD         As String
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'���T�j���Œ�̏j�����
Private Type FixWN
    lMonth      As Long
    lNthWeek    As Long
    lDayOfWeek  As Long
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'�����Œ�̉�Ћx�����
Private Type FixMDC
    sMD         As String
    lDays       As Long
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'���T�j���Œ�̉�Ћx�����
Private Type FixWNC
    lMonth      As Long
    lNthWeek    As Long
    lDayOfWeek  As Long
    lDays       As Long
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'//////////////////////////////////////////////////
'                   �萔
'//////////////////////////////////////////////////

'�u�����̏j���Ɋւ���@���v�{�s�N����
Private Const BEGIN_DATE    As Date = #7/20/1948#

'�u�U�֋x���v�{�s�N����
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE    As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE    As Date = #1/1/2007#

'�u�����̋x���v�{�s�N����
Private Const NATIONAL_HOLIDAY_BEGIN_DATE       As Date = #12/27/1985#

'�N���
Private Const YEAR_MAX      As Long = 2050

'�G���[�R�[�h�i�p�����[�^�ُ�j
Private Const ERROR_INVALID_PARAMETER   As Long = &H57


'//////////////////////////////////////////////////
'               Private�ϐ�
'//////////////////////////////////////////////////

'�����̏j���i�[�p�f�B�N�V���i��
'�L�[�F�N�����iDateTime�^�j
'�l�@�F�j����
Private dicHoliday_ As New Dictionary

Private lInitializedLastYear_   As Long


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'                       �j�����̒�`
'
'�@��{�I�ȏj�����́A�ȉ��̂Q�̃��\�b�h���Œ�`����B
'�@�@getNationalHolidayInfoMD
'�@�@getNationalHolidayInfoWN
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'�����Œ�̏j����񐶐�
'//////////////////////////////////////////////////
Private Sub getNationalHolidayInfoMD(ByRef uFixMD() As FixMD)

    Dim sFixMD(29)  As String   '�j���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim sResult()   As String
    Dim i           As Long

    '//////////////////////////////////////////////////
    '               �����Œ�̏j��
    '//////////////////////////////////////////////////
    '�K�p�J�n�N�ɂ���
    '�@���U�i1/1�j
    '�@���l�̓��i1/15�j
    '�@�V�c�a�����i4/29�j
    '�@���@�L�O���i5/3�j
    '�@���ǂ��̓��i5/5�j
    '�̂T�́A�u�����̏j���Ɋւ���@���v�{�s�N�i1948�N�j�ɐ��肳��Ă��邪
    '���@�̎{�s��7/20�ł���A����ȑO�ƂȂ邽�߁A�K�p�J�n�N�𗂔N�i1949�N�j�ɕ␳���Ă���B
    '
    '����,�K�p�J�n�N,�K�p�I���N,���O
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixMD(0) = "01/01,1949,9999,����"          '�K�p�J�n�N�␳�ς�
    sFixMD(1) = "01/15,1949,1999,���l�̓�"      '�K�p�J�n�N�␳�ς�
    sFixMD(2) = "02/11,1967,9999,�����L�O�̓�"
    sFixMD(3) = "02/23,2020,9999,�V�c�a����"    '�K�p�J�n�N�␳�ς�
    sFixMD(4) = "02/24,1989,1989,���a�V�c�̑�r�̗�"
    sFixMD(5) = "04/10,1959,1959,�c���q���m�e���̌����̋V"
    sFixMD(6) = "04/29,1949,1988,�V�c�a����"    '�K�p�J�n�N�␳�ς�
    sFixMD(7) = "04/29,1989,2006,�݂ǂ�̓�"
    sFixMD(8) = "04/29,2007,9999,���a�̓�"
    sFixMD(9) = "05/01,2019,2019,�V�c�̑���"
    sFixMD(10) = "05/03,1949,9999,���@�L�O��"    '�K�p�J�n�N�␳�ς�
    sFixMD(11) = "05/04,2007,9999,�݂ǂ�̓�"
    sFixMD(12) = "05/05,1949,9999,���ǂ��̓�"    '�K�p�J�n�N�␳�ς�
    sFixMD(13) = "06/09,1993,1993,�c���q���m�e���̌����̋V"
    sFixMD(14) = "07/20,1996,2002,�C�̓�"
    sFixMD(15) = "07/22,2021,2021,�C�̓�"
    sFixMD(16) = "07/23,2020,2020,�C�̓�"
    sFixMD(17) = "07/23,2021,2021,�X�|�[�c�̓�"
    sFixMD(18) = "07/24,2020,2020,�X�|�[�c�̓�"
    sFixMD(19) = "08/08,2021,2021,�R�̓�"
    sFixMD(20) = "08/10,2020,2020,�R�̓�"
    sFixMD(21) = "08/11,2016,2019,�R�̓�"
    sFixMD(22) = "08/11,2022,9999,�R�̓�"
    sFixMD(23) = "09/15,1966,2002,�h�V�̓�"
    sFixMD(24) = "10/10,1966,1999,�̈�̓�"
    sFixMD(25) = "10/22,2019,2019,���ʗ琳�a�̋V"
    sFixMD(26) = "11/03,1948,9999,�����̓�"
    sFixMD(27) = "11/12,1990,1990,���ʗ琳�a�̋V"
    sFixMD(28) = "11/23,1948,9999,�ΘJ���ӂ̓�"
    sFixMD(29) = "12/23,1989,2018,�V�c�a����"

    ReDim uFixMD(UBound(sFixMD))

    For i = 0 To UBound(sFixMD)
        sResult = Split(sFixMD(i), ",")

        uFixMD(i).sMD = sResult(0)
        uFixMD(i).lBeginYear = CLng(sResult(1))
        uFixMD(i).lEndYear = CLng(sResult(2))
        uFixMD(i).sName = sResult(3)
    Next i

End Sub

'//////////////////////////////////////////////////
'���T�j���Œ�̏j����񐶐�
'//////////////////////////////////////////////////
Private Sub getNationalHolidayInfoWN(ByRef uFixWN() As FixWN)

    Dim sFixWN(5)   As String   '�j���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim sResult()   As String
    Dim i           As Long

    '//////////////////////////////////////////////////
    '               ���T�j���Œ�̏j��
    '//////////////////////////////////////////////////
    '��,�T,�j��,�K�p�J�n�N,�K�p�I���N,���O
    '�j���F�� 1
    '�@�@�@�� 2
    '�@�@�@�� 3
    '�@�@�@�� 4
    '�@�@�@�� 5
    '�@�@�@�� 6
    '�@�@�@�y 7
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixWN(0) = "01,2,2,2000,9999,���l�̓�"
    sFixWN(1) = "07,3,2,2003,2019,�C�̓�"
    sFixWN(2) = "07,3,2,2022,9999,�C�̓�"
    sFixWN(3) = "09,3,2,2003,9999,�h�V�̓�"
    sFixWN(4) = "10,2,2,2000,2019,�̈�̓�"
    sFixWN(5) = "10,2,2,2022,9999,�X�|�[�c�̓�"

    ReDim uFixWN(UBound(sFixWN))

    For i = 0 To UBound(sFixWN)
        sResult = Split(sFixWN(i), ",")

        uFixWN(i).lMonth = CLng(sResult(0))
        uFixWN(i).lNthWeek = CLng(sResult(1))
        uFixWN(i).lDayOfWeek = CLng(sResult(2))
        uFixWN(i).lBeginYear = CLng(sResult(3))
        uFixWN(i).lEndYear = CLng(sResult(4))
        uFixWN(i).sName = sResult(5)
    Next i

End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                       �j�����̒�`�@�����܂�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'                       ��Ћx�����̒�`
'
'�@��{�I�ȉ�Ћx�����́A�ȉ��̂T�̃��\�b�h���Œ�`����B
'   getCompanyHolidayInfoW
'   getCompanyHolidayInfoMD
'   getCompanyHolidayInfoWN
'   getCompanyHolidayInfoMDExclude
'   getCompanyHolidayInfoWNExclude
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'�j���Œ�̉�Ћx����񐶐�
'//////////////////////////////////////////////////
Private Sub getCompanyHolidayInfoW(ByRef isHolidayArry() As Boolean)

    Dim isHolidayW(6)   As Boolean
    Dim i               As Long

    isHolidayW(0) = True        '���j��
    isHolidayW(1) = False       '���j��
    isHolidayW(2) = False       '�Ηj��
    isHolidayW(3) = False       '���j��
    isHolidayW(4) = False       '�ؗj��
    isHolidayW(5) = False       '���j��
    isHolidayW(6) = True        '�y�j��

    ReDim isHolidayArry(UBound(isHolidayW))

    For i = 0 To UBound(isHolidayW)
        isHolidayArry(i) = isHolidayW(i)
    Next i

End Sub

'//////////////////////////////////////////////////
'�����Œ�̉�Ћx����񐶐�
'//////////////////////////////////////////////////
Private Sub getCompanyHolidayInfoMD(ByRef uFixMDC() As FixMDC)

    Dim sFixMD(3)  As String    '��Ћx���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim sResult()   As String
    Dim i           As Long

'----�Ώۃf�[�^���Ȃ��ꍇ�AReDim�݂̂�Exit���邱��----
    ReDim uFixMDC(0)
    Exit Sub
'-----------------------------------------------------

    '//////////////////////////////////////////////////
    '               �����Œ�̉�Ћx��
    '//////////////////////////////////////////////////
    '�K�p�J�n����,�K�p����,�K�p�J�n�N,�K�p�I���N,���O
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixMD(0) = "01/02,2,1949,9999,�N�n�x��"
    sFixMD(1) = "09/01,1,1949,9999,�n�ƋL�O��"
    sFixMD(2) = "12/29,3,1949,9999,�N���x��"
    sFixMD(3) = "08/12,7,1949,2018,���~�x��"

    ReDim uFixMDC(UBound(sFixMD))

    For i = 0 To UBound(sFixMD)
        sResult = Split(sFixMD(i), ",")

        uFixMDC(i).sMD = sResult(0)
        uFixMDC(i).lDays = CLng(sResult(1))
        uFixMDC(i).lBeginYear = CLng(sResult(2))
        uFixMDC(i).lEndYear = CLng(sResult(3))
        uFixMDC(i).sName = sResult(4)
    Next i

End Sub

'//////////////////////////////////////////////////
'���T�j���Œ�̉�Ћx����񐶐�
'//////////////////////////////////////////////////
Private Sub getCompanyHolidayInfoWN(ByRef uFixWNC() As FixWNC)

    Dim sFixWNC(1)  As String   '��Ћx���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim sResult()   As String
    Dim i           As Long

'----�Ώۃf�[�^���Ȃ��ꍇ�AReDim�݂̂�Exit���邱��----
    ReDim uFixWNC(0)
    Exit Sub
'-----------------------------------------------------

    '//////////////////////////////////////////////////
    '               ���T�j���Œ�̉�Ћx��
    '//////////////////////////////////////////////////
    '��,�T,�j��,�K�p����,�K�p�J�n�N,�K�p�I���N,���O
    '�j���F�� 1
    '�@�@�@�� 2
    '�@�@�@�� 3
    '�@�@�@�� 4
    '�@�@�@�� 5
    '�@�@�@�� 6
    '�@�@�@�y 7
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixWNC(0) = "06,3,2,1,2000,9999,���ʋx��"
    sFixWNC(1) = "10,4,4,2,2000,9999,���ʋx��"

    ReDim uFixWNC(UBound(sFixWNC))

    For i = 0 To UBound(sFixWNC)
        sResult = Split(sFixWNC(i), ",")

        uFixWNC(i).lMonth = CLng(sResult(0))
        uFixWNC(i).lNthWeek = CLng(sResult(1))
        uFixWNC(i).lDayOfWeek = CLng(sResult(2))
        uFixWNC(i).lDays = CLng(sResult(3))
        uFixWNC(i).lBeginYear = CLng(sResult(4))
        uFixWNC(i).lEndYear = CLng(sResult(5))
        uFixWNC(i).sName = sResult(6)
    Next i

End Sub

'//////////////////////////////////////////////////
'�x���ł��o�΂���v�o�Ώ�񐶐�
'       �����Œ�̉�Џo�Γ�
'//////////////////////////////////////////////////
Private Sub getCompanyHolidayInfoMDExclude(ByVal lLastYear As Long, ByRef dtExclude() As Date)

    Dim sFixMDC(0)  As String   '��Џo�Γ��f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim lElements   As Long
    Dim sResult()   As String
    Dim lDays       As Long
    Dim lYearBegin  As Long
    Dim lYearEnd    As Long
    Dim lYear       As Long
    Dim dtBegin     As Date
    Dim i           As Long
    Dim j           As Long

    lElements = 0
    ReDim dtExclude(lElements)

'----�Ώۃf�[�^���Ȃ��ꍇ�A������Exit���邱��----
    Exit Sub
'------------------------------------------------

    '//////////////////////////////////////////////////
    '               �����Œ�̉�Џo�Γ�
    '//////////////////////////////////////////////////
    '�K�p�J�n����,�K�p����,�K�p�J�n�N,�K�p�I���N,Reserve
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixMDC(0) = "04/29,1,2000,9999,"

    For i = 0 To UBound(sFixMDC)
        sResult = Split(sFixMDC(i), ",")

        lDays = CLng(sResult(1))

        lYearBegin = CLng(sResult(2))

        lYearEnd = CLng(sResult(3))
        If lYearEnd = 9999 Then
            lYearEnd = lLastYear
        End If

        For lYear = lYearBegin To lYearEnd
            dtBegin = CDate(CStr(lYear) & "/" & sResult(0))

            For j = 0 To lDays - 1
                ReDim Preserve dtExclude(lElements)
                dtExclude(lElements) = DateAdd("d", j, dtBegin)

                lElements = lElements + 1
            Next j
        Next lYear
    Next i

End Sub

'//////////////////////////////////////////////////
'�x���ł��o�΂���v�o�Ώ�񐶐�
'       ���T�j���Œ�̉�Џo�Γ�
'//////////////////////////////////////////////////
Private Sub getCompanyHolidayInfoWNExclude(ByVal lLastYear As Long, ByRef dtExclude() As Date)

    Dim sFixWNC(0)  As String   '��Џo�Γ��f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��
    Dim lElements   As Long
    Dim sResult()   As String
    Dim lDays       As Long
    Dim lYearBegin  As Long
    Dim lYearEnd    As Long
    Dim lYear       As Long
    Dim dtBegin     As Date
    Dim i           As Long
    Dim j           As Long

    lElements = 0
    ReDim dtExclude(lElements)

'----�Ώۃf�[�^���Ȃ��ꍇ�A������Exit���邱��----
    Exit Sub
'-----------------------------------------------------

    '//////////////////////////////////////////////////
    '               ���T�j���Œ�̉�Џo�Γ�
    '//////////////////////////////////////////////////
    '��,�T,�j��,�K�p����,�K�p�J�n�N,�K�p�I���N,Reserve
    '�j���F�� 1
    '�@�@�@�� 2
    '�@�@�@�� 3
    '�@�@�@�� 4
    '�@�@�@�� 5
    '�@�@�@�� 6
    '�@�@�@�y 7
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixWNC(0) = "08,4,7,1,2000,9999,"

    ReDim uFixWNC(UBound(sFixWNC))

    For i = 0 To UBound(sFixWNC)
        sResult = Split(sFixWNC(i), ",")

        lDays = CLng(sResult(3))

        lYearBegin = CLng(sResult(4))

        lYearEnd = CLng(sResult(5))
        If lYearEnd = 9999 Then
            lYearEnd = lLastYear
        End If

        For lYear = lYearBegin To lYearEnd
            dtBegin = getNthWeeksDayOfWeek(CStr(lYear), CLng(sResult(0)), CLng(sResult(1)), CLng(sResult(2)))

            For j = 0 To lDays - 1
                ReDim Preserve dtExclude(lElements)
                dtExclude(lElements) = DateAdd("d", j, dtBegin)

                lElements = lElements + 1
            Next j
        Next lYear
    Next i

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                       ��Ћx�����̒�`�@�����܂�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Class_Initialize()

    Set dicHoliday_ = New Dictionary

    lInitializedLastYear_ = &H80000000

    '�f�t�H���g�ŁA���݂̂T�N��܂Ńf�[�^�𐶐�����
    InitializedLastYear = Year(Now) + 5

End Sub

Private Sub Class_Terminate()

    Set dicHoliday_ = Nothing

End Sub


'//////////////////////////////////////////////////
'�w�������Ћx�����H
'//////////////////////////////////////////////////
Public Function isCompanyHoliday(ByVal dtDate As Date) As Boolean

    Dim dtDateW As Date

    '�����b�f�[�^��؂�̂Ă�
    dtDateW = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate))

    If dtDateW < BEGIN_DATE Then
        Err.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", Format$(dtDateW, "yyyy/mm/dd") & "�́A�K�p�͈͊O�ł��B"

        Exit Function
    ElseIf Year(dtDateW) > YEAR_MAX Then
        Err.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", Format$(YEAR_MAX + 1, "yyyy�N") & "�ȍ~�́A�K�p�͈͊O�ł��B"

        Exit Function
    ElseIf Year(dtDateW) > InitializedLastYear Then
        Err.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", Format$(dtDateW, "yyyy�N") & "�́A�f�[�^����������Ă��Ȃ����߁A����ł��܂���B" _
                            & vbCrLf & "reInitialize���\�b�h�őΏ۔N��ݒ��A�ēx�m�F���Ă݂ĉ������B"

        Exit Function
    End If

    isCompanyHoliday = dicHoliday_.Exists(dtDateW)

End Function

'//////////////////////////////////////////////////
'�w�������Ћx�����H�����ł���΁A���̋x���������킹�ĕԂ�
'//////////////////////////////////////////////////
Public Function isCompanyHoliday2(ByVal dtDate As Date, ByRef sHolidayName As String) As Boolean

    Dim dtDateW As Date

    '�����b�f�[�^��؂�̂Ă�
    dtDateW = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate))

    isCompanyHoliday2 = isCompanyHoliday(dtDateW)

    sHolidayName = getCompanyHolidayName(dtDateW)

End Function

'//////////////////////////////////////////////////
'�w��N�̉�Ћx����z��Ɋi�[���ĕԂ�
'//////////////////////////////////////////////////
Public Function getCompanyHolidays(ByVal lYear As Long, ByRef dtHolidays() As Date) As Long

    Dim dtHolidaysW()   As Date
    Dim lHolidays       As Long
    Dim i As Long

    lHolidays = 0
    ReDim dtHolidaysW(lHolidays)

    For i = 0 To dicHoliday_.Count - 1
        If Year(dicHoliday_.Keys(i)) = lYear Then
            ReDim Preserve dtHolidaysW(lHolidays)

            dtHolidaysW(lHolidays) = dicHoliday_.Keys(i)

            lHolidays = lHolidays + 1
        End If
    Next i

    '�������בւ�
    Call qSort(dtHolidaysW, 0, UBound(dtHolidaysW))

    Erase dtHolidays
    dtHolidays = dtHolidaysW

    getCompanyHolidays = lHolidays

End Function

'//////////////////////////////////////////////////
'�w����̉�Ћx������Ԃ�
'//////////////////////////////////////////////////
Public Function getCompanyHolidayName(ByVal dtHoliday As Date) As String

    Dim dtDateW As Date

    '�����b�f�[�^��؂�̂Ă�
    dtDateW = DateSerial(Year(dtHoliday), Month(dtHoliday), Day(dtHoliday))

    If isCompanyHoliday(dtDateW) = True Then
        getCompanyHolidayName = dicHoliday_.Item(dtDateW)
    End If

End Function

'//////////////////////////////////////////////////
'���N�܂ł̉�Ћx���f�[�^����������Ă��邩
'//////////////////////////////////////////////////
Public Property Get InitializedLastYear() As Long

    InitializedLastYear = lInitializedLastYear_

End Property

'//////////////////////////////////////////////////
'�w��N�܂ł̉�Ћx���f�[�^�𐶐�������iYEAR_MAX�ȉ��j
'�@�O������̗v���́AreInitialize�ōs�����Ƃ��o����
'//////////////////////////////////////////////////
Private Property Let InitializedLastYear(ByVal lInitializedLastYear As Long)

    If lInitializedLastYear < lInitializedLastYear_ Then
        '�v�����ꂽ�ŏI�N���������ς݂̔N���O�Ȃ�΁A�������Ȃ�
        Exit Property
    ElseIf lInitializedLastYear > YEAR_MAX Then
        lInitializedLastYear = YEAR_MAX
    End If

    Call initDictionary(lInitializedLastYear)

    lInitializedLastYear_ = lInitializedLastYear

End Property

'//////////////////////////////////////////////////
'�w��N�܂ł̉�Ћx���f�[�^�𐶐�������
'//////////////////////////////////////////////////
Public Sub reInitialize(ByVal lLastYear As Long)

    InitializedLastYear = lLastYear

End Sub

'//////////////////////////////////////////////////
'Dictionary�։�Ћx�������i�[
'//////////////////////////////////////////////////
Private Sub initDictionary(ByVal lLastYear As Long)

    Dim uFixMD()    As FixMD
    Dim uFixWN()    As FixWN
    Dim isHolidayArry() As Boolean
    Dim uFixMDC()   As FixMDC
    Dim uFixWNC()   As FixWNC
    Dim dtExclude() As Date

    '�����Œ�̏j�����
    Call getNationalHolidayInfoMD(uFixMD)

    '���T�j���Œ�̏j�����
    Call getNationalHolidayInfoWN(uFixWN)

    'Dictionary�֒ǉ�
    Call add2Dictionary(lLastYear, uFixMD, uFixWN)


    '�j���Œ�̉�Ћx�����
    Call getCompanyHolidayInfoW(isHolidayArry)

    '�����Œ�̉�Ћx�����
    Call getCompanyHolidayInfoMD(uFixMDC)

    '���T�j���Œ�̉�Ћx�����
    Call getCompanyHolidayInfoWN(uFixWNC)

    'Dictionary�֒ǉ�
    Call add2DictionaryC(lLastYear, uFixMDC, uFixWNC, isHolidayArry)


    '�j���ł��o�΂���v�o�Ώ��i�����Œ�j
    Call getCompanyHolidayInfoMDExclude(lLastYear, dtExclude)

    'Dictionary����v�o�Γ������O
    Call removeFromDictionaryC(dtExclude)

    ReDim dtExclude(0)

    '�j���ł��o�΂���v�o�Ώ��i���T�j���Œ�j
    Call getCompanyHolidayInfoWNExclude(lLastYear, dtExclude)

    'Dictionary����v�o�Γ������O
    Call removeFromDictionaryC(dtExclude)

End Sub

'//////////////////////////////////////////////////
'�j������Dictionary�֊i�[
'//////////////////////////////////////////////////
Private Sub add2Dictionary(ByVal lLastYear As Long, ByRef uFixMD() As FixMD, ByRef uFixWN() As FixWN)

    Dim lInitializedLastYear    As Long
    Dim lBeginYear          As Long
    Dim lEndYear            As Long
    Dim dtHoliday           As Date
    Dim lAddedDays          As Long
    Dim dtAdded()           As Date
    Dim existsHoliday       As Boolean
    Dim lYear               As Long
    Dim i                   As Long

    '�������ς݂̍ŏI�N���擾
    lInitializedLastYear = InitializedLastYear

    If lInitializedLastYear < Year(BEGIN_DATE) Then
        '�{�H�N���O�Ȃ�΁A�{�H�N���J�n�N�Ƃ���
        lBeginYear = Year(BEGIN_DATE)
    Else
        '�{�H�N�Ȍ�Ȃ�A�������ς݂̗��N���J�n�N�Ƃ���
        lBeginYear = lInitializedLastYear + 1
    End If

    lEndYear = lLastYear

    For lYear = lBeginYear To lEndYear
        '�N�Ԃ̏j���i�[�p�z��N���A
        lAddedDays = 0
        ReDim dtAdded(lAddedDays)

        '�����Œ�̏j��
        For i = 0 To UBound(uFixMD)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If uFixMD(i).lBeginYear <= lYear And uFixMD(i).lEndYear >= lYear Then
                dtHoliday = CDate(CStr(lYear) & "/" & uFixMD(i).sMD)

                dicHoliday_.Add dtHoliday, uFixMD(i).sName

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '���T�j���Œ�̏j��
        For i = 0 To UBound(uFixWN)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If uFixWN(i).lBeginYear <= lYear And uFixWN(i).lEndYear >= lYear Then
                dtHoliday = getNthWeeksDayOfWeek(lYear, uFixWN(i).lMonth, uFixWN(i).lNthWeek, uFixWN(i).lDayOfWeek)

                dicHoliday_.Add dtHoliday, uFixWN(i).sName

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '�t���̓�
        dtHoliday = getVernalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "�t���̓�"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '�H���̓�
        dtHoliday = getAutumnalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "�H���̓�"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '�U�֋x��
        For i = 0 To lAddedDays - 1
            existsHoliday = existsSubstituteHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "�U�֋x��"
            End If
        Next i

        '�����̋x��
        For i = 0 To lAddedDays - 1
            existsHoliday = existsNationalHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "�����̋x��"
            End If
        Next i

        Erase dtAdded
    Next lYear

End Sub

'//////////////////////////////////////////////////
'��Ћx������Dictionary�֊i�[
'//////////////////////////////////////////////////
Private Sub add2DictionaryC(ByVal lLastYear As Long, ByRef uFixMDC() As FixMDC, ByRef uFixWNC() As FixWNC, ByRef isHolidayArry() As Boolean)

    Dim lInitializedLastYear    As Long
    Dim lBeginYear          As Long
    Dim lEndYear            As Long
    Dim dtHolidayBegin      As Date
    Dim dtHoliday           As Date
    Dim lYear               As Long
    Dim dtBeginDate         As Date
    Dim lDays               As Long
    Dim dtTargetDate        As Date
    Dim i                   As Long
    Dim j                   As Long

    '�������ς݂̍ŏI�N���擾
    lInitializedLastYear = InitializedLastYear

    If lInitializedLastYear < Year(BEGIN_DATE) Then
        '�{�H�N���O�Ȃ�΁A�{�H�N���J�n�N�Ƃ���
        lBeginYear = Year(BEGIN_DATE)
    Else
        '�{�H�N�Ȍ�Ȃ�A�������ς݂̗��N���J�n�N�Ƃ���
        lBeginYear = lInitializedLastYear + 1
    End If

    lEndYear = lLastYear

    For lYear = lBeginYear To lEndYear
        '�����Œ�̉�Ћx��
        If uFixMDC(0).sMD Like "##/##" Then
            For i = 0 To UBound(uFixMDC)
                '�K�p���Ԃ݂̂�ΏۂƂ���
                If uFixMDC(i).lBeginYear <= lYear And uFixMDC(i).lEndYear >= lYear Then
                    dtHolidayBegin = CDate(CStr(lYear) & "/" & uFixMDC(i).sMD)

                    For j = 0 To uFixMDC(i).lDays - 1
                        dtHoliday = DateAdd("d", j, dtHolidayBegin)

                        'Dictionary�ɖ��o�^�̏ꍇ�̂ݒǉ�����
                        If dicHoliday_.Exists(dtHoliday) = False Then
                            dicHoliday_.Add dtHoliday, uFixMDC(i).sName
                        End If
                    Next j
                End If
            Next i
        End If

        '���T�j���Œ�̉�Ћx��
        If uFixWNC(0).lMonth >= 1 And uFixWNC(0).lMonth <= 12 Then
            For i = 0 To UBound(uFixWNC)
                '�K�p���Ԃ݂̂�ΏۂƂ���
                If uFixWNC(i).lBeginYear <= lYear And uFixWNC(i).lEndYear >= lYear Then
                    dtHolidayBegin = getNthWeeksDayOfWeek(lYear, uFixWNC(i).lMonth, uFixWNC(i).lNthWeek, uFixWNC(i).lDayOfWeek)

                    For j = 0 To uFixWNC(i).lDays - 1
                        dtHoliday = DateAdd("d", j, dtHolidayBegin)

                        'Dictionary�ɖ��o�^�̏ꍇ�̂ݒǉ�����
                        If dicHoliday_.Exists(dtHoliday) = False Then
                            dicHoliday_.Add dtHoliday, uFixWNC(i).sName
                        End If
                    Next j
                End If
            Next i
        End If

        '�e�T�Œ�̉�Ћx��
        dtBeginDate = DateSerial(lYear, 1, 1)
        lDays = DateDiff("d", dtBeginDate, DateSerial(lYear, 12, 31))

        For i = 0 To lDays
            dtTargetDate = DateAdd("d", i, dtBeginDate)

            If isHolidayArry(Weekday(dtTargetDate) - vbSunday) Then
                If dicHoliday_.Exists(dtTargetDate) = False Then
                    dicHoliday_.Add dtTargetDate, "��Ћx��"
                End If
            End If
        Next i
    Next lYear

End Sub

'//////////////////////////////////////////////////
'��Ћx������Dictionary���珜�O
'//////////////////////////////////////////////////
Private Sub removeFromDictionaryC(ByRef dtExclude() As Date)

    Dim i   As Long

    If IsDate(dtExclude(0)) = False Then
        Exit Sub
    End If

    For i = 0 To UBound(dtExclude)
        'Dictionary�ɓo�^�ς̏ꍇ���O����
        If dicHoliday_.Exists(dtExclude(i)) Then
            dicHoliday_.Remove dtExclude(i)
        End If
    Next i

End Sub

'//////////////////////////////////////////////////
'�U�֋x���̗L��
'�@�j���idtDate�j�ɑ΂���U�֋x���̗L���i����ꍇ�́AdtSubstituteHoliday�ɑ�������j
'//////////////////////////////////////////////////
Private Function existsSubstituteHoliday(ByVal dtDate As Date, ByRef dtSubstituteHoliday As Date) As Boolean

    Dim dtNextDay   As Date

    existsSubstituteHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDate���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If dtDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And dtDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(dtDate) = vbSunday Then
            '�j�������j���ł���΁A�����i���j���j���U�֋x��
            dtSubstituteHoliday = DateAdd("d", 1, dtDate)

            existsSubstituteHoliday = True
        End If
    ElseIf dtDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '�u�����̏j���v�����j���ɓ�����Ƃ��́A���̓���ɂ����Ă��̓��ɍł��߂��u�����̏j���v�łȂ������x���Ƃ���
        If Weekday(dtDate) = vbSunday Then
            dtNextDay = DateAdd("d", 1, dtDate)

            '���߂̏j���łȂ������擾
            Do Until dicHoliday_.Exists(dtNextDay) = False
                dtNextDay = DateAdd("d", 1, dtNextDay)
            Loop

            dtSubstituteHoliday = dtNextDay

            existsSubstituteHoliday = True
        End If
    End If

End Function

'//////////////////////////////////////////////////
'�����̋x���̗L��
'�@�j���idtDate�j�ɑ΂������̋x���̗L���i����ꍇ�́AdtNationalHoliday�ɑ�������j
'//////////////////////////////////////////////////
Private Function existsNationalHoliday(ByVal dtDate As Date, ByRef dtNationalHoliday As Date) As Boolean

    Dim dtBaseDay   As Date
    Dim dtNextDay   As Date

    existsNationalHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDate���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If dtDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        dtBaseDay = DateAdd("d", 1, dtDate)

        '���߂̏j���łȂ������擾
        Do Until dicHoliday_.Exists(dtBaseDay) = False
            dtBaseDay = DateAdd("d", 1, dtBaseDay)
        Loop

        '���j���ł���ΑΏۊO
        If Weekday(dtBaseDay) <> vbSunday Then
            dtNextDay = DateAdd("d", 1, dtBaseDay)

            '�������j���ł���ΑΏ�
            If dicHoliday_.Exists(dtNextDay) = True Then
                existsNationalHoliday = True

                dtNationalHoliday = dtBaseDay
            End If
        End If
    End If

End Function

'//////////////////////////////////////////////////
'���̑�N W�j���̓������擾
'//////////////////////////////////////////////////
Private Function getNthWeeksDayOfWeek(ByVal lYear As Long, _
                                      ByVal lMonth As Long, _
                                      ByVal lNth As Long, _
                                      ByVal lDayOfWeek As VbDayOfWeek) As Date

    Dim dt1stDate       As Date
    Dim lDayOfWeek1st   As Long
    Dim lOffset         As Long

    '�w��N���̂P�����擾
    dt1stDate = DateSerial(lYear, lMonth, 1)

    '�P���̗j�����擾
    lDayOfWeek1st = Weekday(dt1stDate)

    '�w����ւ̃I�t�Z�b�g���擾
    lOffset = lDayOfWeek - lDayOfWeek1st

    If lDayOfWeek1st > lDayOfWeek Then
        lOffset = lOffset + 7
    End If

    lOffset = lOffset + 7 * (lNth - 1)

    getNthWeeksDayOfWeek = DateAdd("d", lOffset, dt1stDate)

End Function

'//////////////////////////////////////////////////
'�t���̓����擾
'//////////////////////////////////////////////////
Private Function getVernalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(20.8431 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getVernalEquinoxDay = DateSerial(lYear, 3, lDay)

End Function

'//////////////////////////////////////////////////
'�H���̓����擾
'//////////////////////////////////////////////////
Private Function getAutumnalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(23.2488 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getAutumnalEquinoxDay = DateSerial(lYear, 9, lDay)

End Function

Private Sub qSort(ByRef dtHolidays() As Date, ByVal lLeft As Long, ByVal lRight As Long)

    Dim dtCenter    As Date
    Dim dtTemp      As Date
    Dim i           As Long
    Dim j           As Long

    If lLeft < lRight Then
        dtCenter = dtHolidays((lLeft + lRight) \ 2)

        i = lLeft - 1
        j = lRight + 1

        Do While (True)
            i = i + 1
            Do While (dtHolidays(i) < dtCenter)
                i = i + 1
            Loop

            j = j - 1
            Do While (dtHolidays(j) > dtCenter)
                j = j - 1
            Loop

            If i >= j Then
                Exit Do
            End If

            dtTemp = dtHolidays(i)
            dtHolidays(i) = dtHolidays(j)
            dtHolidays(j) = dtTemp
        Loop

        Call qSort(dtHolidays, lLeft, i - 1)
        Call qSort(dtHolidays, j + 1, lRight)
    End If

End Sub