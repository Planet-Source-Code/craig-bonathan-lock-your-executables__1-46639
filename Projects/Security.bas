Attribute VB_Name = "Security"
Public Password As String

Const AppSize As Long = 24576

Function ReadApp()
    Dim FileNum As Integer
    Dim PassLength As Byte
    Dim FileLength As Long
    Dim FileData As String
    FileNum = FreeFile
    
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #FileNum
        Get #FileNum, AppSize + 1, PassLength
        Password = String(PassLength, Chr(0))
        Get #FileNum, AppSize + 2, Password
        FileLength = LOF(FileNum) - 2 - PassLength
        FileData = String(FileLength, Chr(0))
        Get #FileNum, AppSize + 2 + PassLength, FileData
    Close #FileNum
    
    FileNum = FreeFile
    Open App.Path & "\TempExe01234.exe" For Binary As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
End Function

Function ReadPassword()
    Dim FileNum As Integer
    Dim PassLength As Byte
    FileNum = FreeFile
    
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #FileNum
        Get #FileNum, AppSize + 1, PassLength
        Password = String(PassLength, Chr(0))
        Get #FileNum, AppSize + 2, Password
    Close #FileNum
End Function
