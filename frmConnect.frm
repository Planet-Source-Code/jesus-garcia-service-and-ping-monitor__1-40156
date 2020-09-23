VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Computer"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin Service.FlatButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin Service.FlatButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Ok"
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
   Begin VB.TextBox txtComputerName 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   345
      Width           =   4065
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer Name"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   90
      Width           =   3450
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeepName As String


Private Sub cmdCancel_Click()
' If txtComputerName.Text = "" Then
'    Tag = ""
' Else
'    Tag = txtComputerName
' End If
    Tag = KeepName
    Hide
End Sub

Private Sub cmdOK_Click()
    Tag = txtComputerName
    Hide
End Sub

Private Sub Form_Activate()
    txtComputerName.SelStart = 0
    txtComputerName.SelLength = Len(txtComputerName.Text)
    txtComputerName.SetFocus
End Sub

Private Sub Form_Load()
    If txtComputerName.Text = "" Then
        txtComputerName.Text = LocalName
    End If
    KeepName = txtComputerName.Text
End Sub

Private Sub txtComputerName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
