VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Executable"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "?"
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Timer CheckForApp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   0
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Caption         =   "Authorization Not Provided"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your authorization code below:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckForApp_Timer()
    If CheckIfOpen = False Then
        Kill App.Path & "\TempExe01234.exe"
        End
    End If
End Sub

Private Sub Command1_Click()
    If Text1.Text = Password Then
        Status.Caption = "Authorization Code Valid" & Chr(10) & Chr(13) & "Access Allowed"
        Status.ForeColor = RGB(0, 128, 0)
        Text1.Text = ""
        Text1.Visible = False
        Command1.Visible = False
        Label1.Caption = "Please Wait..."
        ReadApp
        Main.Hide
        OpenApplication (App.Path & "\TempExe01234.exe")
        CheckForApp.Enabled = True
    Else
        Status.Caption = "Authorization Code Invalid" & Chr(10) & Chr(13) & "Access Denied"
        Status.ForeColor = RGB(128, 0, 0)
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ReadPassword
End Sub
