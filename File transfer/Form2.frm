VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About Odesa File Transfer"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   9000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   0
      X2              =   9000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   4440
      X2              =   4440
      Y1              =   3000
      Y2              =   4680
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ODESA FILE TRANSFER (WINSOCK)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Programmed And  By : Alper ESKIKILIC E-Mail: odesayazilim@gmail.com "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   240
      Picture         =   "Form2.frx":55B9
      Top             =   3600
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   6000
      Picture         =   "Form2.frx":5FA3
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "www.odesayazilim.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   21
      Left            =   5400
      MouseIcon       =   "Form2.frx":9922
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4320
      Width           =   3465
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label_Click(Index As Integer)
ExecLink "http://www.odesayazilim.com", vbNormalFocus
End Sub

Private Sub ExecLink(Url As String, style As VbAppWinStyle)
Shell "explorer.exe " & Url & "", stlye
End Sub

