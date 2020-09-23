VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkForce 
      Caption         =   "Force"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox chkLowPower 
      Caption         =   "&Prevent computer from entering Stand By or Hibernating"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1000
      Width           =   3015
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "&Shutdown"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cbShutdown 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Power As New CPower

Private Sub chkLowPower_Click()
    Power.DisableSuspend = CBool(chkLowPower.Value)
End Sub

Private Sub cmdShutdown_Click()
    With cbShutdown
    Select Case .Text
        Case "Shut Down"
            Power.ShutDown lShutDown, CBool(chkForce.Value)
        Case "Reboot"
            Power.ShutDown lReboot, CBool(chkForce.Value)
        Case "Log Off"
            Power.ShutDown lLogOff, CBool(chkForce.Value)
        Case "Stand by"
            Power.ShutDown lSuspend, CBool(chkForce.Value)
        Case "Hibernate"
            Power.ShutDown lHibernate, CBool(chkForce.Value)
    End Select: End With
    
    
End Sub

Private Sub Form_Load()
    With cbShutdown
        .AddItem "Shut Down"
        .AddItem "Reboot"
        .AddItem "Log Off"
        .AddItem "Stand by"
        If CanHibernate Then .AddItem "Hibernate"
        .ListIndex = 0
    End With
    
    Power.Initialize Me.hwnd
End Sub
