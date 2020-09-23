VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1605
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   1605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMenu 
      Caption         =   "Menu"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin Service.FlatButton cmdServiceMonitor 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Service Monitor"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
      End
      Begin Service.FlatButton cmdPingMonitor 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Ping Monitor"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
      End
      Begin Service.FlatButton cmdQuit 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Quit"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSM 
         Caption         =   "Service Monitor"
      End
      Begin VB.Menu mnuPM 
         Caption         =   "Ping Monitor"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim moSysTray As CSysTray
Private Sub cdmExitApp_Click()
    Unload Me
    End
End Sub

Private Sub cmdPing_Click()
    frmPing.Show vbModal, Me
End Sub

Private Sub cmdServices_Click()
    frmService.Show vbModal, Me
End Sub

Private Sub flbPingMonitor_CLICKED()
    frmPing.Show vbModal, Me
End Sub

Private Sub flbQuit_CLICKED()
    Unload Me
    End
End Sub

Private Sub flbServiceMonitor_CLICKED()
    frmService.Show vbModal, Me
End Sub

Private Sub cmdPingMonitor_Click()
    frmPing.Show vbModal, Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
    End
End Sub

Private Sub cmdServiceMonitor_Click()
    frmService.Show vbModal, Me
End Sub

Private Sub FlatButton1_Click()

End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        MsgBox "Look at your sys tray, the program is running...", vbOKOnly + vbInformation, "Advise"
        Unload Me
        End
    End If
    mnuMain.Visible = False
    'Read Service.ini
    Call ReadINI
    'Create object for Systray
    Set moSysTray = New CSysTray
    moSysTray.ToolTipText = "Service and Ping Monitor"
    Set moSysTray.Client = Me
    WindowState = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moSysTray = Nothing
End Sub

Private Sub mnuCancel_Click()
    mnuMain.Visible = False
End Sub

Private Sub mnuPM_Click()
    frmPing.Show vbModal, Me
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    End
End Sub

Private Sub mnuSM_Click()
    frmService.Show vbModal, Me
End Sub
