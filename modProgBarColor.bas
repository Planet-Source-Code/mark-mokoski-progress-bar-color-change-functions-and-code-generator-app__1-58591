Attribute VB_Name = "modProgBarColor"
    '********************************************************
    '
    '   Progress Bar Functions
    '   31-JAN-2005
    '
    '   Mark Mokoski
    '   28-OCT-2003
    '   markm@cmtelephone.com
    '   http://www.rjillc.com
    '
    '   Set of functions to change colors of the stock
    '   (and plain) MS Common Controls Progress Bar.
    '   The fuctions are from a user control I did in 2003 for myself.
    '   But I found that I only used these functions most of the time,
    '   just changing the Foreground and Background colors.
    '   So why add "Code Bloat" with a user control when a few functions
    '   will do the job.
    '
    '**********************************************************
    
    Option Explicit
    
    Private Const WM_USER                    As Long = &H400
    Private Const CCM_FIRST                  As Long = &H2000
    Private Const CCM_SETBKCOLOR             As Long = (CCM_FIRST + 1)
    Private Const PBM_SETBARCOLOR            As Long = (WM_USER + 9)
    Private Const PBM_SETBKCOLOR             As Long = CCM_SETBKCOLOR
    
    Public Const CLR_DEFAULT                 As Long = &HFF000000
    
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, lparam As Any) As Long

    
Public Function pbForeColor(ByRef pbControl As Control, ByVal pbColor As Long)

    'Set the Progress Bar ForeColor
    SendMessage pbControl.hwnd, PBM_SETBARCOLOR, 0, ByVal pbColor

End Function

Public Function pbBackColor(ByRef pbControl As Control, ByVal pbColor As Long)

    'Set the Progress Bar Backcolor
    SendMessage pbControl.hwnd, PBM_SETBKCOLOR, 0, ByVal pbColor

End Function

Public Function pbDefaultColor(ByRef pbControl As Control)

    'Set the Progress Bar to default colors
    SendMessage pbControl.hwnd, PBM_SETBARCOLOR, 0, ByVal CLR_DEFAULT
    SendMessage pbControl.hwnd, PBM_SETBKCOLOR, 0, ByVal CLR_DEFAULT
    
End Function
