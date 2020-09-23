VERSION 5.00
Begin VB.Form frmPing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping Monitor"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraIp 
      Caption         =   "Ip to monitor"
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      Begin VB.ListBox lstPing 
         Height          =   2205
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
      Begin Service.FlatButton cmdSave 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Save"
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
      Begin Service.FlatButton cmdEditIP 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Edit"
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
      Begin Service.FlatButton cmdRemoveIP 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Remove"
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
   Begin Service.FlatButton cmdAddIP 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "Add IP"
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
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Timer tmrPing 
      Interval        =   60000
      Left            =   3480
      Top             =   120
   End
   Begin Service.FlatButton cmdExitIP 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Exit"
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
      Caption         =   "Ip:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   180
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ECHO As ICMP_ECHO_REPLY
Dim iMin As Integer

Private Sub cmdAddIp_Click()
    If txtIP.Text <> "" Then
        lstPing.AddItem txtIP.Text
        txtIP.Text = ""
    End If
    txtIP.SetFocus
End Sub

Private Sub cmdCacel_Click()
    Call cmdExitIP_Click
End Sub

Private Sub cmdCancel_Click()
    cmdExitIP_Click
End Sub

Private Sub cmdEditIP_Click()
    MsgBox "Method not finished yet...", vbOKOnly + vbInformation, Me.Caption
End Sub

Private Sub cmdExitIP_Click()
    Me.Hide
End Sub

Private Sub cmdRemoveIp_Click()
    lstPing.RemoveItem lstPing.ListIndex
End Sub

Private Sub cmdSave_Click()
Dim Ff As Integer
Dim Buf As String

    Ff = FreeFile
    Open App.Path & "\IpAddress.Txt" For Output As Ff
        For i = 0 To lstPing.ListCount - 1
            Buf = lstPing.List(i)
            Print #Ff, Buf
        Next
    Close Ff
End Sub

Private Sub cmdTest_Click()
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub FlatButton1_Click()

End Sub

Private Sub Form_Load()
Dim Ff As Integer
Dim Buf As String

    Ff = FreeFile
    Open App.Path & "\IpAddress.Txt" For Input As Ff
        While Not EOF(Ff)
            Input #Ff, Buf
            lstPing.AddItem Buf
        Wend
    Close Ff
    iMin = 0
End Sub


Private Sub tmrPing_Timer()
Dim i As Integer
Dim j As Integer
Dim CurIP As String
Dim iTimes As Byte
Dim stHour As String
    
    iMin = iMin + 1
    If iMin >= iTimer Then 'iTimer defined en ReadINI sub on module modCode
        iMin = 0
        For i = 0 To lstPing.ListCount - 1
            lstPing.Selected(i) = True
            CurIP = lstPing.Text
            iTimes = 0
            For j = 1 To iPingTimes 'iPingTimes defined en ReadINI sub on module modCode
                Wait (500)
                Call Ping(CurIP, ECHO)
                Select Case ECHO.status
                    Case 11010
                        iTimes = iTimes + 1
                    Case Else
                End Select
            Next j
            If iTimes >= iPingsFailed Then 'iPingsFailed  defined en ReadINI sub on module modCode
                stHour = Format(Now(), "HHMM")
                If stHour <= "1800" Then
                    For j = LBound(arComputers) To UBound(arComputers)
                        Shell "Net Send " & arComputers(j) & " The IP address: " & CurIP & " failed at pinging." & vbCrLf & _
                              iTimes & " of " & iPingTimes & "times." & _
                              vbCrLf & vbCrLf & _
                              "This ping test was made at " & Format(Now(), "dddd, dd of mmmm at HH:MM:SS") & _
                              vbCrLf & _
                              "Note: This colud happend because of the traffic." & _
                              vbCrLf & vbCrLf & vbCrLf & _
                              "Copyright ALMEX 2002", vbHide
                    Next j
                Else
                    For j = LBound(arComputers) To UBound(arComputers) 'Vector Read from INI
                        If UCase(arComputers(j)) <> "MTY_FBELTRAN" Then
                            Shell "Net Send " & arComputers(j) & " The IP address: " & CurIP & " failed at pinging." & vbCrLf & _
                                  iTimes & " of " & iPingTimes & "times." & _
                                  vbCrLf & vbCrLf & _
                                  "This ping test was made at " & Format(Now(), "dddd, dd of mmmm at HH:MM:SS") & _
                                  vbCrLf & _
                                  "Note: This colud happend because of the traffic." & _
                                  vbCrLf & vbCrLf & vbCrLf & _
                                  "Copyright ALMEX 2002", vbHide
                        End If
                    Next j
                End If
            End If
        Next i
        iMin = 0
    End If
End Sub
