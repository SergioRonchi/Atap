Attribute VB_Name = "mdlNetworkname"
Option Explicit
Declare Function GetUserName Lib "advapi32.dll" _
    Alias "GetUserNameA" (ByVal lpBuffer As String, _
    nSize As Long) As Long

Public Declare Function GetComputerName Lib "Kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function ComputerName() As String
    Dim sBuffer As String * 255
    If GetComputerName(sBuffer, 255&) <> 0 Then
        ComputerName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        ComputerName = "Sconosciuto"
    End If
End Function

Public Function UserName() As String
    Dim lpBuff As String * 25
    Dim ret As Long
    'Get the user name minus any trailing
    'spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = UCase(Left(lpBuff, InStr(lpBuff, Chr(0)) - 1))
   
End Function



