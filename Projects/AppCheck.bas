Attribute VB_Name = "AppCheck"
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public AppIdentity As Long

Function OpenApplication(FileName As String) As Boolean
    AppIdentity = Shell(FileName, vbNormalFocus)
    If AppIdentity > 0 Then
        OpenApplication = True
    Else
        OpenApplication = False
    End If
End Function

Function CheckIfOpen() As Boolean
    Dim Temp As Long
    Temp = OpenProcess(1, 0, AppIdentity)
    If Temp > 0 Then
        CheckIfOpen = True
        Temp = CloseHandle(Temp)
    Else
        CheckIfOpen = False
    End If
End Function
