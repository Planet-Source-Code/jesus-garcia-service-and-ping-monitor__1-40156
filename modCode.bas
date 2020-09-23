Attribute VB_Name = "modCode"
Declare Function WNetGetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount Lib "KERNEL32" () As Long

'-----------------------------------------------
'---------Public Constants----------------------
'------------------------------------------------

'--- Defined in winnt.h ----------------

Public Const SERVICE_BOOT_START = &H0
Public Const SERVICE_SYSTEM_START = &H1
Public Const SERVICE_AUTO_START = &H2
Public Const SERVICE_DEMAND_START = &H3
Public Const SERVICE_DISABLED = &H4

'
' Error control type
'
Public Const SERVICE_ERROR_IGNORE = &H0
Public Const SERVICE_ERROR_NORMAL = &H1
Public Const SERVICE_ERROR_SEVERE = &H2
Public Const SERVICE_ERROR_CRITICAL = &H3

Public Const SERVICE_KERNEL_DRIVER = &H1
Public Const SERVICE_FILE_SYSTEM_DRIVER = &H2
Public Const SERVICE_ADAPTER = &H4
Public Const SERVICE_RECOGNIZER_DRIVER = &H8

Public Const SERVICE_DRIVER = &HB
Public Const SERVICE_WIN32_OWN_PROCESS = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS = &H20
Public Const SERVICE_WIN32 = &H30
Public Const SERVICE_INTERACTIVE_PROCESS = &H100


Public Const ADS_SERVICE_STOPPED = &H1
Public Const ADS_SERVICE_START_PENDING = &H2
Public Const ADS_SERVICE_STOP_PENDING = &H3
Public Const ADS_SERVICE_RUNNING = &H4
Public Const ADS_SERVICE_CONTINUE_PENDING = &H5
Public Const ADS_SERVICE_PAUSE_PENDING = &H6
Public Const ADS_SERVICE_PAUSED = &H7
Public Const ADS_SERVICE_ERROR = &H8
Public LocalName As String
Public arComputers() As String
Public iPingTimes As Integer
Public iPingsFailed As Integer
Public iTimer As Integer

Public Function Get_ComputerName()
    Dim lpBuff As String * 25
    Dim ret As Long, ComputerName As String
    
    ret = GetComputerName(lpBuff, 25)
    ComputerName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

    Get_ComputerName = ComputerName
End Function

Public Sub ReadINI()
Dim oIni As clsINI
Dim stComputers As String
Dim stPingTimes As String
Dim stPingsFailed As String
Dim stTimer As String

    Set oIni = New clsINI
    oIni.Path = App.Path & "\Service.ini"
    stComputers = oIni.GetValue("General", "Computers", "MTY_Chuy")
    stPingTimes = oIni.GetValue("General", "PingTimes", "10")
    stPingsFailed = oIni.GetValue("General", "PingsFailed", "10")
    stTimer = oIni.GetValue("General", "TimerMin", "5")
    arComputers = Split(stComputers, ",")
    iPingTimes = Val(stPingTimes)
    iPingsFailed = Val(stPingsFailed)
    iTimer = Val(stTimer)
    For i = LBound(arComputers) To UBound(arComputers)
        arComputers(i) = Trim(arComputers(i))
    Next
    Set oIni = Nothing
End Sub
Public Sub Wait(ByVal dblMilliseconds As Double)
Dim dblStart As Double
Dim dblEnd As Double
Dim dblTickCount As Double

    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
End Sub

