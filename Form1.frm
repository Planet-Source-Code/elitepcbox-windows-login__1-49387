VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WINDOWS LOGIN"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDomain 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "DOMAIN:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "LOGIN: " & VerifyLogin(txtUser.Text, txtDomain.Text, txtPass.Text)
End Sub

Private Sub Form_Load()
    Dim x As Long
    Dim strName As String
    strName = String$(255, 0)
    x = GetUserName(strName, Len(strName))
    txtUser.Text = strName
End Sub
