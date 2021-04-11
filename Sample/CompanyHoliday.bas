Attribute VB_Name = "CompanyHoliday"
Option Explicit

Public Sub ‹x“úˆê——Žæ“¾()

    Dim cch     As CCompanyHoliday
    Dim dt()    As Date
    Dim i       As Long

    Set cch = New CCompanyHoliday

    Call cch.getCompanyHolidays(2021, dt)

    For i = 0 To UBound(dt)
        Debug.Print dt(i), cch.getCompanyHolidayName(dt(i))
    Next i

End Sub
