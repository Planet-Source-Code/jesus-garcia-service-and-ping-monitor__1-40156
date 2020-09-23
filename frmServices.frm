VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Monitor"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Icon            =   "frmServices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Service.FlatButton cmdClose 
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Hide Monitor"
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
   Begin VB.Timer tmrAnimation 
      Interval        =   5000
      Left            =   6600
      Top             =   600
   End
   Begin MSComctlLib.ImageList imglstClock 
      Left            =   5400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":0992
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":0CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":0FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":12E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":15FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":1914
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":1C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":1F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":2262
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":257C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":2896
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServices.frx":2BB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   4560
   End
   Begin VB.Timer tmrMon 
      Interval        =   60000
      Left            =   5400
      Top             =   4560
   End
   Begin VB.Frame fraServices 
      Caption         =   "&Services"
      Height          =   3615
      Left            =   120
      TabIndex        =   23
      Top             =   1050
      Width           =   2295
      Begin VB.ListBox lstService 
         Height          =   2790
         Left            =   120
         TabIndex        =   24
         Top             =   285
         Width           =   2085
      End
      Begin Service.FlatButton cmdRefresh 
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Refresh status"
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
   Begin VB.Frame fraLog 
      Caption         =   "Monitor Log"
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   6720
      Width           =   6855
      Begin VB.ListBox lstLog 
         Height          =   1230
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame fraMonitor 
      Caption         =   "&Monitored Services"
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   6855
      Begin VB.ListBox lstMonApp 
         Height          =   450
         Left            =   5400
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ListBox lstMon 
         Height          =   1425
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5175
      End
      Begin Service.FlatButton cmdAdd 
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Add Monitor"
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
      Begin Service.FlatButton cmdRemove 
         Height          =   375
         Left            =   5400
         TabIndex        =   32
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Remove Monitor"
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   8370
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSvcStatus 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Top             =   3240
      Width           =   2925
   End
   Begin VB.TextBox txtSvcType 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   2055
      Width           =   2925
   End
   Begin VB.TextBox txtOrderGroup 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      Top             =   2865
      Width           =   2925
   End
   Begin VB.TextBox txtErrorControl 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   2460
      Width           =   2925
   End
   Begin VB.TextBox txtStartType 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   1650
      Width           =   2925
   End
   Begin VB.TextBox txtComputer 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   795
      TabIndex        =   1
      Top             =   255
      Width           =   2745
   End
   Begin VB.TextBox txtDisplayName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   1275
      Width           =   2925
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   2745
      Left            =   2535
      TabIndex        =   6
      Top             =   1050
      Width           =   4470
      Begin VB.Label Label5 
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Order Group:"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1965
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Error Control:"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lblSvcType 
         Caption         =   "Service Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1155
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Startup:"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   810
      End
      Begin VB.Label DisplayName 
         Caption         =   "Display Name:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Width           =   1485
      End
   End
   Begin VB.Frame fraOperations 
      Caption         =   "Operations"
      Height          =   735
      Left            =   2505
      TabIndex        =   12
      Top             =   3945
      Width           =   4500
      Begin VB.Timer Timer2 
         Interval        =   30000
         Left            =   4200
         Top             =   240
      End
      Begin Service.FlatButton cmdStart 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Start"
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
      Begin Service.FlatButton cmdStop 
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Stop"
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
      Begin Service.FlatButton cmdPause 
         Height          =   375
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Pause"
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
   Begin VB.Frame Frame3 
      Caption         =   "Computer "
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   15
      Width           =   5160
      Begin Service.FlatButton cmdChange 
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Change Computer"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.Image ImgClock 
      Height          =   480
      Left            =   6000
      Top             =   555
      Width           =   480
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim cont As IADsContainer
Dim iCurrentImage As Integer


Private Sub cmdAdd_Click()
Dim svc As IADsService
Dim svcOps As IADsServiceOperations
Dim i As Integer
Dim bolExist As Boolean
    
    bolExist = False
    For i = 0 To lstMonApp.ListCount - 1
        If lstService.Text = lstMonApp.List(i) Then
            bolExist = True
            Exit For
        End If
    Next
    If Not bolExist Then
        Set svc = GetCurrentService()
        lstMon.AddItem svc.DisplayName
        lstMonApp.AddItem lstService.Text
        Set svc = Nothing
    Else
        MsgBox "This service is already monitored...", vbOKOnly + vbCritical, "Error"
    End If
    lstService.SetFocus
    tmrMon.Enabled = True
    iCurrentImage = 1
    tmrAnimation.Enabled = True
    ImgClock.Visible = True
    ImgClock.Picture = imglstClock.ListImages(1).Picture
    
End Sub

Private Sub cmdChange_Click()
    ShowComputerDlg
    PopulateService (txtComputer)
End Sub

Private Sub ShowComputerDlg()
    frmConnect.Show vbModal, Me
    If (frmConnect.Tag = "") Then
        txtComputer = "LocalHost"
    Else
        txtComputer = frmConnect.Tag
    End If
End Sub
Private Sub PopulateService(computerStr As String)
Dim RespYesNo As Integer
    On Error Resume Next
    lstService.Clear
    Set cont = GetObject("WinNT://" & computerStr & ",computer")
    If Err.Number = -2147024843 Then 'The computer doesn't exist
        RespYesNo = MsgBox("The computer " & computerStr & " doesn't exist." & vbCrLf & _
                           "Will use " & LocalName & " instead. Are you agree?", vbYesNo + vbCritical, "Advise")
        If RespYesNo = vbYes Then
            txtComputer.Text = LocalName
            computerStr = LocalName
            Set cont = GetObject("WinNT://" & computerStr & ",computer")
        Else
            cmdChange_Click
        End If
    End If
    cont.Filter = Array("Service")
    For Each svc In cont
        lstService.AddItem svc.Name
    Next
End Sub

Private Function GetCurrentService() As IADsService
 If (lstService.Text = "") Then
    Set GetCurrentService = Nothing
    Exit Function
 End If
    
 Set GetCurrentService = cont.GetObject("service", lstService.Text)

End Function

Private Function GetCurrentServiceMon() As IADsService
 If (lstMonApp.Text = "") Then
    Set GetCurrentServiceMon = Nothing
    Exit Function
 End If
    
 Set GetCurrentServiceMon = cont.GetObject("service", lstMonApp.Text)

End Function


Private Sub cmdClose_Click()
    frmService.Hide
End Sub

Private Sub cmdPause_Click()
Dim svc As IADsService
Dim svcOp As IADsServiceOperations

    
Set svc = GetCurrentService()
Set svcOp = svc
On Error GoTo ErrorHandler
svcOp.Pause

'Refresh the status by simulating user's selection
lstService_Click
Timer2.Enabled = True
Timer2_Timer

Exit Sub
ErrorHandler:
If Err.Number = -2147023844 Then
MsgBox "This option is not available for this service type", vbExclamation, "This Option Not Available"
Else
MsgBox Err.Number & ":  " & Err.Description, vbExclamation, "This Option Unavailable"
End If
End Sub

Private Sub cmdRefresh_Click()
 lstService_Click
End Sub

Private Sub cmdRemove_Click()
Dim IndexMon As Integer
    
    IndexMon = lstMon.ListIndex
    If IndexMon >= 0 Then
        lstMon.RemoveItem (IndexMon)
        lstMonApp.RemoveItem (IndexMon)
    End If
    lstMon.SetFocus
    If lstMon.ListCount - 1 >= 0 Then
        If IndexMon < lstMon.ListCount - 1 Then
            lstMon.Selected(IndexMon) = True
        Else
            If IndexMon <> 0 Then
                lstMon.Selected(IndexMon - 1) = True
            Else
                lstMon.Selected(0) = True
            End If
        End If
    End If
    If lstMon.ListCount < 1 Then
        tmrMon.Enabled = False
        tmrAnimation.Enabled = False
        ImgClock.Visible = False
        iCurrentImage = 1
    End If
End Sub

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim svc As IADsService
Dim svcOp As IADsServiceOperations

    
Set svc = GetCurrentService()
Set svcOp = svc

svcOp.Start

'Refresh the status by simulating user's selection
lstService_Click
Timer2.Enabled = True
Timer2_Timer
Exit Sub

ErrorHandler:
MsgBox Err.Number & ":  " & Err.Description, vbExclamation, "This Option Unavailable"
End Sub

Private Sub cmdStop_Click()
On Error GoTo ErrorHandler
Dim svc As IADsService
Dim svcOp As IADsServiceOperations

    
Set svc = GetCurrentService()
Set svcOp = svc

svcOp.Stop

'Refresh the status by simulating user's selection
lstService_Click
Timer2.Enabled = True
Timer2_Timer
Exit Sub

ErrorHandler:
MsgBox Err.Number & ":  " & Err.Description, vbExclamation, "This Option Unavailable"

End Sub

Private Sub Command1_Click()
End Sub

Private Sub FlatButton1_Click()

End Sub

Private Sub Form_Activate()
    lstService.SetFocus
    lstService.Selected(0) = True
End Sub

Private Sub Form_Load()
 
    tmrMon.Enabled = False
    tmrAnimation.Enabled = False
    ImgClock.Visible = False
    iCurrentImage = 1
    'Get local computer name
    LocalName = Get_ComputerName
    txtComputer.Text = LocalName
    PopulateService (txtComputer.Text)
    StatusBar1.Panels(1) = Now()
    Timer1.Enabled = True
    Timer1_Timer
    StatusBar1.Panels(2) = txtComputer.Text
    StatusBar1.Panels(4) = "User -> " & Get_User_Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    frmService.Hide
End Sub

Private Sub lstMon_Click()
    lstMonApp.Selected(lstMon.ListIndex) = True
    'MsgBox lstMon.ListIndex
End Sub

Private Sub lstMon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Call cmdRemove_Click
    End If
End Sub

Private Sub lstMonApp_Click()
    lstMon.Selected(lstMonApp.ListIndex) = True
End Sub

Private Sub lstService_Click()
Dim svc As IADsService
Dim svcOps As IADsServiceOperations
    
Set svc = GetCurrentService()

txtDisplayName = svc.DisplayName

'Order Group
txtOrderGroup = svc.LoadOrderGroup

Select Case svc.StartType
  Case SERVICE_BOOT_START
    txtStartType = "Boot Start"
  Case SERVICE_SYSTEM_START
    txtStartType = "System Start"
  Case SERVICE_AUTO_START
    txtStartType = "Automatic"
  Case SERVICE_DEMAND_START
    txtStartType = "Manual"
  Case SERVICE_DISABLED
    txtStartType = "Disabled"
  Case Else
    txtStartType = "Unknown"
End Select

'Error Control
Select Case svc.ErrorControl

Case SERVICE_ERROR_IGNORE
   txtErrorControl = "Service ignores error"
Case SERVICE_ERROR_NORMAL
   txtErrorControl = "No Error"
Case SERVICE_ERROR_SEVERE
 txtErrorControl = "Severe error"
Case SERVICE_ERROR_CRITICAL
 txtErrorControl = "Critical error"
Case Else
  txtErrorControl = "Unknown"
End Select

'-------------------------------------------
' Service Type
'---------------------------------------------
Select Case svc.ServiceType

Case SERVICE_KERNEL_DRIVER
     txtSvcType = "Kernel Driver"
Case SERVICE_FILE_SYSTEM_DRIVER
     txtSvcType = "File System Driver"
Case SERVICE_ADAPTER
     txtSvcType = "Adapter"
Case SERVICE_RECOGNIZER_DRIVER
     txtSvcType = "Recognizer Driver"
Case SERVICE_WIN32_OWN_PROCESS
     txtSvcType = "Win32 Process"
Case SERVICE_WIN32_SHARE_PROCESS
     txtSvcType = "Win32 Share Process"
Case SERVICE_WIN32
     txtSvcType = "Win32"
Case SERVICE_INTERACTIVE_PROCESS
     txtSvcType = "Interactive Process"
End Select
                                        
'--- Get the Service Status
Set svcOps = svc

Select Case svcOps.status
Case ADS_SERVICE_STOPPED:
   txtSvcStatus = "Stopped"
Case ADS_SERVICE_START_PENDING:
   txtSvcStatus = "Start Pending"
Case ADS_SERVICE_STOP_PENDING:
   txtSvcStatus = "Stop Pending"
Case ADS_SERVICE_RUNNING:
   txtSvcStatus = "Running"
Case ADS_SERVICE_CONTINUE_PENDING:
  txtSvcStatus = "Continue Pending"
Case ADS_SERVICE_PAUSE_PENDING:
 txtSvcStatus = "Pause Pending"
Case ADS_SERVICE_PAUSED:
 txtSvcStatus = "Paused"
Case ADS_SERVICE_ERROR:
 txtSvcStatus = "Error"
End Select

Set svc = Nothing
Set svcOps = Nothing
End Sub

Public Function Get_User_Name()
    Dim s$, cnt&, dl&
    cnt& = 199
    s$ = String$(200, 0)
    dl& = WNetGetUserName(s$, cnt)
    Get_User_Name = Left$(s$, cnt)
End Function

Private Sub Timer1_Timer()
    StatusBar1.Panels(1) = Now()
End Sub

Private Sub Timer2_Timer()
    'MsgBox "About to check services..."
    lstService_Click
End Sub


Private Sub tmrAnimation_Timer()

    If iCurrentImage < 12 Then
        iCurrentImage = iCurrentImage + 1
    Else
        iCurrentImage = 1
    End If
    ImgClock.Picture = imglstClock.ListImages(iCurrentImage).Picture
    ImgClock.Refresh
End Sub

Private Sub tmrMon_Timer()
Dim ServiceName As String
Dim i As Integer
Dim j As Integer
Dim svc As IADsService
Dim svcOp As IADsServiceOperations
Dim svcOps As IADsServiceOperations
Dim Ff As Integer
Dim LogSt As String
    
    StatusBar1.Panels(3) = "Quering services..."
    StatusBar1.Refresh
    For i = 0 To lstMonApp.ListCount - 1
        lstMonApp.Selected(i) = True
        ServiceName = lstMonApp.Text
        Set svc = GetCurrentServiceMon()
        Set svcOp = svc
        Set svcOps = svc
        Select Case svcOps.status
        Case ADS_SERVICE_STOPPED:
            svcOp.Start
            LogSt = "Service " & Mid(lstMon.List(i) & Space(40), 1, 40) & " restarted on " & txtComputer.Text & " at " & Now
            lstLog.AddItem LogSt
            lstLog.Refresh
            Ff = FreeFile
            Open App.Path & "\" & txtComputer.Text & ".Log" For Output As Ff
                Write #Ff, LogSt
            Close Ff
            For j = LBound(arComputers) To UBound(arComputers)
                Shell "Net Send " & arComputers(j) & " " & "Service " & lstMon.List(i) & " restarted on " & txtComputer.Text & "...", vbHide
            Next
            'Shell "Net Send MTY_Juan03" & " " & "Service " & lstMon.List(i) & " restarted on " & txtComputer.Text & "...", vbHide
        End Select
        Set svc = Nothing
        Set svcOp = Nothing
        Set svcOps = Nothing
        DoEvents
    Next
    StatusBar1.Panels(3) = ""
    StatusBar1.Refresh
End Sub
