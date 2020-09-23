VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
' Module Name:      clsINI
' Module Type:      Class Module
' Filename:         clsINI.cls
' Author:           Dino Damalas
' Date:             11/5/2001
' Purpose:          quick class for working with INI files - this really didn't
'                   have to be a class, but what the hell....
'
'*****************************************************************************
Option Explicit

Private m_sPath As String

'api declarations
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Property Let Path(sPath As String)
    ' The filename of the INI file:
    m_sPath = sPath
End Property


Property Get Path() As String
    ' The filename of the INI file:
    Path = m_sPath
End Property


'===============================================================================
' Name:     GetValue
' Input:
'           strSection  - section in the INI
'           strKey      - key name for value to be retrievied
'           strDefault  - default value
' Output:
'           string - value of matching key in INI file
' Purpose:  <Description (Abstract)>
' Remarks:  <Things to take care of>
' Author:   Dino Damalas
' Version:  <(optional for v1.0) internal version number>
' Date:     11/5/2001
'===============================================================================
Public Function GetValue(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strDefault As String) As String
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(strSection, strKey, strDefault, sBuf, iSize, m_sPath)


    If (iSize > 0) Then
        GetValue = Left$(sBuf, iRetCode)
    Else
        GetValue = ""
    End If
End Function

'===============================================================================
' Name:     SetValue
' Input:
'           strSection  - section in the INI
'           strKey      - key name for value to be retrievied
'           strValue    - value to be savied
' Output:
'           long - result of save
' Purpose:  <Description (Abstract)>
' Remarks:  <Things to take care of>
' Author:   Dino Damalas
' Version:  <(optional for v1.0) internal version number>
' Date:     11/5/2001
'===============================================================================
Public Function SetValue(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String) As Long

    Dim iPos As Integer
    ' Strip chr$(0):
    iPos = InStr(strValue, Chr$(0))


    Do While iPos <> 0
        strValue = Left$(strValue, (iPos - 1)) & Mid$(strValue, (iPos + 1))
        iPos = InStr(strValue, Chr$(0))
    Loop
    SetValue = WritePrivateProfileString(strSection, strKey, strValue, m_sPath)
    
End Function
'deletes a value in the INI file if a given section / key
Public Function DeleteValue(ByVal strSection As String, _
                         ByVal strKey As String) As Long
    ' Delete the value at Key within Section
    '     of Path
    DeleteValue = WritePrivateProfileString(strSection, strKey, 0&, m_sPath)
End Function

'deletes a section in teh INI file
Public Function DeleteSection(ByVal strSection As String) As Long
    ' Delete the Section in Path
    DeleteSection = WritePrivateProfileString(strSection, 0&, 0&, m_sPath)
End Function




'===============================================================================
' Name:     subErrorHandler
' Input:
'           objError - error objectc
'           strMethodName - name of method where error occured.
' Output:
'           none
' Purpose:  logs error to the event log
' Remarks:  <Things to take care of>
' Author:   Dino Damalas
' Date:     11/5/2001
'===============================================================================
Private Sub subErrorHandler(ByRef objError As ErrObject, ByVal strMethodName As String)
    
    
    
    Dim intLineNumber As Integer
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strLogDescription As String
    
    On Error Resume Next
    
    lngErrorNumber = objError.Number
    strErrorDescription = objError.Description
    
    strLogDescription = "Application: " & App.Title & ";" & _
                                      "Version: " & App.Major & "." & App.Minor & "." & App.Revision & ";" & _
                                      "Method: " & strMethodName & ";" & _
                                      "Error Number: " & lngErrorNumber & ";" & _
                                      "Error Description: " & strErrorDescription & ";"
                                    
                        
    Debug.Print strLogDescription
    
    Call App.LogEvent(strLogDescription, vbLogEventTypeError)
           
    
End Sub

