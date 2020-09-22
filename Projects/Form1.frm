VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Executable"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Password"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Combine"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "File to protect"
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "LockExe.exe"
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim FileNum As Integer
    Dim FileData1 As String
    Dim FileData2 As String
    Dim PassLength As Byte
    Dim Password As String
    
    FileNum = FreeFile
    Open App.Path & "\" & Text1.Text For Binary Access Read As #FileNum
        FileData1 = String(LOF(FileNum), Chr(0))
        Get #FileNum, 1, FileData1
    Close #FileNum
    
    FileNum = FreeFile
    Open App.Path & "\" & Text2.Text For Binary Access Read As #FileNum
        FileData2 = String(LOF(FileNum), Chr(0))
        Get #FileNum, 1, FileData2
    Close #FileNum
    
    
    Password = Text3.Text
    PassLength = Len(Password)
    
    FileData1 = FileData1 & Chr(PassLength) & Password & FileData2
    
    FileNum = FreeFile
    Open App.Path & "\Locked012345.exe" For Binary As #FileNum
        Put #FileNum, 1, FileData1
    Close #FileNum
End Sub
