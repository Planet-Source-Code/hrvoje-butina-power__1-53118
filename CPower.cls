VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim lHwnd As Long

Public Property Let DisableSuspend(ByVal SuspendDisabled As Boolean)
    bPreventLowPower = SuspendDisabled
    If SuspendDisabled Then
        If oldProc = 0 Then oldProc = SetWindowLong(lHwnd, GWL_WNDPROC, AddressOf WindowProc)
    Else
        If Not oldProc = 0 Then SetWindowLong lHwnd, GWL_WNDPROC, oldProc
    End If
    
End Property

Public Property Get DisableSuspend() As Boolean
    DisableSuspend = bPreventLowPower
End Property

Public Sub Initialize(ByVal hwnd As Long)
    lHwnd = hwnd
End Sub

Public Sub ShutDown(ShutDownType As eShutDownType, Optional Force As Boolean = True)
    If ShutDownType = lHibernate Or ShutDownType = lSuspend Then
        LowPowerState ShutDownType - 5, Force
    Else
        ShutDownW ShutDownType, Force
    End If
End Sub

Private Sub Class_Terminate()
    If Not oldProc = 0 Then SetWindowLong lHwnd, GWL_WNDPROC, oldProc
End Sub
