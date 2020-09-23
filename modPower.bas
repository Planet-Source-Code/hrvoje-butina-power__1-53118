Attribute VB_Name = "modPower"
'Module for handling all those neat stuff in this program
'Note that you'll have to change few thins if you want
'to use this module (without class) alone, like
'subclass when program starts etc.
Option Explicit

Private Const BROADCAST_QUERY_DENY = &H424D5144
Public Const GWL_WNDPROC = -4
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const PWR_HIBERNATE = 5
Private Const PWR_SUSPEND = 6
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const ANYSIZE_ARRAY = 1
Private Const PBT_APMQUERYSUSPEND = &H0
Private Const PBT_APMRESUMESUSPEND = &H7
Private Const WM_POWERBROADCAST = &H218

Public Enum eShutDownType
    lShutDown = EWX_SHUTDOWN
    lReboot = EWX_REBOOT
    lLogOff = EWX_LOGOFF
    lHibernate = PWR_HIBERNATE
    lSuspend = PWR_SUSPEND
End Enum

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type LUID
    LowPart As Long
    HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Private Declare Function SetSystemPowerState Lib "kernel32.dll" (ByVal fSuspend As Long, ByVal bForce As Boolean) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Public oldProc As Long
Public bPreventLowPower As Boolean

Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

'AdjustTokenPrivileges for Winnt
Private Sub EnableShutDown()
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
End Sub

Public Sub ShutDownW(iShutdownType As eShutDownType, bForce As Boolean)
    'Simply shutdowns computer
    Dim lFlags As Long
    lFlags = iShutdownType
    If bForce Then lFlags = lFlags + EWX_FORCE
    If IsWinNT Then EnableShutDown
    ExitWindowsEx lFlags, 0
End Sub

Public Sub LowPowerState(iPower As Integer, bForce As Boolean)
    '0 - Hibernate : 1 - Suspend
    'if there is no Hibernate support then it will suspend
    If IsWinNT Then EnableShutDown
    SetSystemPowerState iPower, bForce
End Sub

Public Function CanHibernate() As Boolean
    'Simply checks if there is hiberfil.sys file
    If Not Dir$("c:\hiberfil.sys", vbHidden + vbReadOnly + vbSystem) = vbNullString Then CanHibernate = True
End Function


'All windows recieve WM_POWERBROADCAST message
'when windows is going to hibernate or suspend
'with wParam set PBT event

'here are just examples of PBT events
'Private Const PBT_APMBATTERYLOW = &H9
'Private Const PBT_APMPOWERSTATUSCHANGE = &HA
'Private Const PBT_APMRESUMESUSPEND = &H7
'Private Const PBT_APMSUSPEND = &H4

'This can be very useful : for example you need to know
'when computer goes standby so you can prepare
'your program for standby (clean variables, pause some procedures etc.)
Public Function WindowProc(ByVal hwnd As Long, ByVal message As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

    'If window recieves Suspend message
    If message = WM_POWERBROADCAST And lparam = PBT_APMQUERYSUSPEND And bPreventLowPower Then
        WindowProc = BROADCAST_QUERY_DENY ' Deny request to enter low power
        Debug.Print "Goin' low power... "
    Else
        WindowProc = CallWindowProc(oldProc, hwnd, message, wparam, lparam) ' Continue
    End If
End Function


