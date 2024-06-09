Attribute VB_Name = "modFileTime"
'===============================================================================
'   Constants
'===============================================================================

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

'===============================================================================
'   Types
'===============================================================================
Private Type SYSTEMTIME
    wYear                   As Integer
    wMonth                  As Integer
    wDayOfWeek              As Integer
    wDay                    As Integer
    wHour                   As Integer
    wMinute                 As Integer
    wSecond                 As Integer
    wMilliseconds           As Integer
End Type

Public Type FILETIME                            ' Win32 date
    dwLowDateTime           As Long
    dwHighDateTime          As Long
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                    As Long
    StandardName(63)        As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(63)        As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type

'===============================================================================
'   Private Members
'===============================================================================

'===============================================================================
'   Declares
'===============================================================================
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Public Function UtcToLocalFileTime(FILETIME As FILETIME) As FILETIME
'===============================================================================
'   UtcToLocalFileTime - Converts local FILETIME to UTC/GMT FILETIME.
'===============================================================================

    Dim Success     As Boolean

    ' Exit if null date supplied.
    If FILETIME.dwHighDateTime = 0 _
    And FILETIME.dwLowDateTime = 0 Then
        Exit Function
    End If

    ' Convert to local time
    Success = FileTimeToLocalFileTime(FILETIME, UtcToLocalFileTime)
    Debug.Assert Success

End Function

Public Function UtcFromLocalFileTime(FILETIME As FILETIME) As FILETIME
'===============================================================================
'   UtcFromLocalFileTime - Converts UTC/GMT FILETIME to a local FILETIME.
'===============================================================================

    Dim Success     As Boolean

    ' Exit if null date supplied.
    If FILETIME.dwHighDateTime = 0 _
    And FILETIME.dwLowDateTime = 0 Then
        Exit Function
    End If

    ' Convert to UTC time
    Success = LocalFileTimeToFileTime(FILETIME, UtcFromLocalFileTime)
    Debug.Assert Success

End Function

Public Function FileTimeToDate(FILETIME As FILETIME) As Date
'===============================================================================
'   FileTimeToDate - Converts FILETIME structure to a VB Date data type.
'
'   NOTE: The FILETIME structure is a structure of 100-nanosecond intervals since
'   January 1, 1601.  The VB Date data type is a floating point value where the
'   value to the left of the decimal is the number of days since December 30,
'   1899, and the value to the right of the decimal represents the time. The
'   hour of the time value can be calculated by multiplying by 24, with the
'   remainder multiplied by 60 to get the minutes, and that remainder can then
'   be multiplied by 60 to get the seconds.
'
'
'   FileTime            The FILETIME structure to convert.
'
'   RETURNS             A date/time value in the intrinsic VB Date data type.
'
'===============================================================================

    Dim Success     As Boolean
    Dim SysTime     As SYSTEMTIME

    Success = FileTimeToSystemTime(FILETIME, SysTime)
    'Debug.Assert Success
    
    ' Create a date/time value from the system time parts
    With SysTime
        FileTimeToDate = _
            DateSerial(.wYear, .wMonth, .wDay) + _
            TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function FileTimeFromDate(FromDate As Date) As FILETIME
'===============================================================================
'   FileTimeFromDate - Converts a VB Date data type to a FILETIME structure.
'
'   NOTE: The FILETIME structure is a structure of 100-nanosecond intervals since
'   January 1, 1601.  The VB Date data type is a floating point value where the
'   value to the left of the decimal is the number of days since December 30,
'   1899, and the value to the right of the decimal represents the time. The
'   hour of the time value can be calculated by multiplying by 24, with the
'   remainder multiplied by 60 to get the minutes, and that remainder can then
'   be multiplied by 60 to get the seconds.
'
'
'   FromDate            The VB DAte to convert.
'
'   RETURNS             A date/time value in the native Win32 FILETIME structure.
'
'===============================================================================

    Dim Success     As Boolean
    Dim SysTime     As SYSTEMTIME

    ' Create SYSTEMTIME from each date part in a date/time value.
    With SysTime
        .wYear = Year(FromDate)
        .wMonth = Month(FromDate)
        .wDay = Day(FromDate)
        .wHour = Hour(FromDate)
        .wMinute = Minute(FromDate)
        .wSecond = Second(FromDate)
    End With

    ' convert the SYSTEMTIME to the FILETIME
    Success = SystemTimeToFileTime(SysTime, FileTimeFromDate)
    Debug.Assert Success

End Function

Public Function UtcToLocalTime(ByVal UtcTime As Date) As Date
'===============================================================================
'   UtcToLocalTime - Converts UTC/GMT time to local time.
'===============================================================================

    Dim FILETIME        As FILETIME

    FILETIME = FileTimeFromDate(UtcTime)
    FILETIME = UtcToLocalFileTime(FILETIME)
    UtcToLocalTime = FileTimeToDate(FILETIME)

End Function

Public Function UtcFromLocalTime(ByVal LocalTime As Date) As Date
'===============================================================================
'   UtcFromLocalTime - Converts local time to UTC/GMT time.
'===============================================================================

    Dim FILETIME        As FILETIME

    FILETIME = FileTimeFromDate(LocalTime)
    FILETIME = UtcFromLocalFileTime(FILETIME)
    UtcFromLocalTime = FileTimeToDate(FILETIME)

End Function


Public Function FileTimeToDate_2(Ft As FILETIME) As Date
' FILETIME units are 100s of nanoseconds.
Const TICKS_PER_SECOND = 10000000

Dim lo_time As Double
Dim hi_time As Double
Dim seconds As Double
Dim hours As Double
Dim the_date As Date
    On Error GoTo ErrHandler

    ' Get the low order data.
    If Ft.dwLowDateTime < 0 Then
        lo_time = 2 ^ 31 + (Ft.dwLowDateTime And &H7FFFFFFF)
    Else
        lo_time = Ft.dwLowDateTime
    End If

    ' Get the high order data.
    If Ft.dwHighDateTime < 0 Then
        hi_time = 2 ^ 31 + (Ft.dwHighDateTime And &H7FFFFFFF)
    Else
        hi_time = Ft.dwHighDateTime
    End If

    ' Combine them and turn the result into hours.
    seconds = (lo_time + 2 ^ 32 * hi_time) / TICKS_PER_SECOND
    hours = CLng(seconds / 3600)
    seconds = seconds - hours * 3600

    ' Make the date.
    the_date = DateAdd("h", hours, "01/01/1601 0:00 AM")
    the_date = DateAdd("s", seconds, the_date)
    FileTimeToDate_2 = the_date
ErrHandler:
    MsgBox Err.Description

End Function
Public Function IsValidDate(sDate As String) As Boolean
On Error GoTo ErrHandler
Dim temp As Date
    temp = CDate(sDate)
    IsValidDate = True
    Exit Function
ErrHandler:
IsValidDate = False

End Function
'
Public Function DateText(Month As String) As Integer
Select Case UCase(Month)
    Case "JAN"
        DateText = 1
    Case "FEB"
        DateText = 2
    Case "MAR"
        DateText = 3
    Case "APR"
        DateText = 4
    Case "MAY"
        DateText = 5
    Case "JUN"
        DateText = 6
    Case "JUL"
        DateText = 7
    Case "AUG"
        DateText = 8
    Case "SEP"
        DateText = 9
    Case "OCT"
        DateText = 10
    Case "NOV"
        DateText = 11
    Case "DEC"
        DateText = 12
    Case Else
        DateText = 1
End Select
    
End Function
