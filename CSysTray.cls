VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'API Declarations
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'API Structures
Private Type NOTIFYICONDATA
    cbSize              As Long
    hWnd                As Long
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

'API constants
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

'Local variable
Private muTrayID As NOTIFYICONDATA
Private WithEvents moClient As Form
Attribute moClient.VB_VarHelpID = -1
Private msToolTipText As String
Private moIcon As StdPicture
Public Property Set Icon(ByVal poIcon As StdPicture)
    Set moIcon = poIcon
End Property
Private Sub Add()
   With muTrayID
      .cbSize = Len(muTrayID)
      .hWnd = moClient.hWnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = IIf(moIcon Is Nothing, moClient.Icon, moIcon)
      .szTip = msToolTipText
   End With
   Call Shell_NotifyIcon(NIM_ADD, muTrayID)
End Sub
Public Property Set Client(poClient As Form)
    Set moClient = poClient
    Call Add
End Property
Private Sub Remove()
    Call Shell_NotifyIcon(NIM_DELETE, muTrayID)
End Sub
Public Property Let ToolTipText(psToolTipText As String)
    msToolTipText = psToolTipText & Chr(0)
End Property

Private Sub Class_Terminate()
    Call Remove
End Sub

Private Sub moClient_Load()
    Call Add
End Sub
Private Sub moClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lResult As Long
   Dim lMessage As Long
   
   If moClient.ScaleMode = vbPixels Then
      lMessage = X
   Else
      lMessage = X / Screen.TwipsPerPixelX
   End If
      
   Select Case lMessage
      Case WM_LBUTTONUP
         moClient.Show
         moClient.WindowState = vbNormal
         lResult = SetForegroundWindow(moClient.hWnd)
      Case WM_RBUTTONUP
         'moClient.Show
         'moClient.WindowState = vbNormal
         'lResult = SetForegroundWindow(moClient.hWnd)
         If moClient.WindowState = vbMinimized Then
            moClient.PopupMenu frmMain.mnuMain
         End If
   End Select
End Sub

Private Sub moClient_Resize()
    If moClient.WindowState = vbMinimized Then moClient.Hide
End Sub

Private Sub moClient_Terminate()
    Call Remove
End Sub
Private Sub moClient_Unload(Cancel As Integer)
    Call Remove
End Sub
