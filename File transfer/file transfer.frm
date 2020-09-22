VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "file transfer"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "file transfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin Project1.CommandXP CommandXP3 
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   1410
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Quit Sys"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":628A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP CommandXP2 
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   1410
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "About Sys"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":62A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP Command4 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   810
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "Send File"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":62C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP Command5 
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "Disconnect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":62DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP Command3 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "Browse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":62FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP Command2 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "Connnect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":6316
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.CommandXP CommandXP1 
      Height          =   300
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      BTYPE           =   1
      TX              =   "Listen"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "file transfer.frx":6332
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   285
      TabIndex        =   10
      Top             =   1440
      Width           =   4785
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7560
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   285
      TabIndex        =   8
      Top             =   855
      Width           =   4785
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   4365
      TabIndex        =   4
      Top             =   270
      Width           =   825
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2655
      TabIndex        =   3
      Top             =   270
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   285
      TabIndex        =   1
      Top             =   270
      Width           =   825
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6960
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Left            =   4350
      Top             =   255
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Left            =   2640
      Top             =   255
      Width           =   1665
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   300
      Left            =   270
      Top             =   255
      Width           =   855
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   255
      Top             =   1410
      Width           =   4830
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   255
      Top             =   840
      Width           =   4830
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "save directory:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   285
      TabIndex        =   11
      Top             =   1215
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "port:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Left            =   4365
      TabIndex        =   9
      Top             =   45
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "file to send:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   285
      TabIndex        =   7
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "No Connection"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "connect to:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   2655
      TabIndex        =   2
      Top             =   45
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "listen on port:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Version
  major As Integer
  minor As Integer
  revision As Integer
End Type

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Dim fileData As String, fileSize As Long, lFileName As String, iWritePos As Long
Dim Sending As Boolean, Receiving As Boolean, sProgress As Long
Dim Lft As Long, Tp As Long, noSave As Boolean, noLoad As Boolean
Dim acceptConnections As Boolean, acceptTransfers As Boolean, autoListen As Boolean

Private Type POINT
  X As Long
  Y As Long
End Type
Dim SutcuImam As POINT, dRagging As Boolean

Private Sub Button1_Click()

End Sub

Private Sub CommandXP1_Click()
On Error Resume Next
If PathIsDirectory(Text1.Text) = 0 Then
  MsgBox "before you can listen or connect, your save directory must be valid", vbOKOnly Or vbCritical, "error"
  Label4.Caption = "invalid save directory"
  Text1.SetFocus
  Exit Sub
End If
If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"

If Text2.Text = "" Then
  MsgBox "please fill in the liston on port box"
  Exit Sub
Else
  If Winsock1.State = 7 Then _
    If MsgBox("you are currently connected. disconnect?", vbYesNo Or vbQuestion, "disconnect?") = vbNo Then Exit Sub
  Winsock1.Close
  Sending = False
  Receiving = False
  Winsock1.LocalPort = Text2.Text
  Winsock1.Listen
  Label4.Caption = "listening on " & Winsock1.LocalIP
  Text1.Enabled = False
End If
End Sub

Private Sub CommandXP2_Click()
Form2.Show
End Sub

Private Sub CommandXP3_Click()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SutcuImam.X = X
SutcuImam.Y = Y
dRagging = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dRagging = True Then
  Form1.Left = Form1.Left + X - SutcuImam.X
  Form1.Top = Form1.Top + Y - SutcuImam.Y
  DoEvents
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
dRagging = False

End Sub


Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()
If PathIsDirectory(Text1.Text) = 0 Then
  MsgBox "before you can listen or connect, your save directory must be valid", vbOKOnly Or vbCritical, "error"
  Label4.Caption = "invalid save directory"
  Text1.SetFocus
  Exit Sub
End If
If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"

If Text3.Text = "" Or Text4.Text = "" Then
  MsgBox "please fill in both boxes"
  Exit Sub
Else
  If Winsock1.State = 7 Then _
    If MsgBox("you are currently connected. disconnect?", vbYesNo Or vbQuestion, "disconnect?") = vbNo Then Exit Sub
  Winsock1.Close
  Winsock1.Connect Text3.Text, Val(Text4.Text)
  Label4.Caption = "connecting..."
  Text1.Enabled = False
End If

End Sub

Private Sub Command3_Click()
On Error Resume Next

CD1.ShowOpen
If Err.Number = cdlCancel Then Exit Sub
Text5.Text = CD1.FileName

End Sub

Private Sub Command4_Click()
If Winsock1.State <> 7 Then
  MsgBox "you are not connected", vbOKOnly Or vbCritical, "error"
  Exit Sub
Else
  If Sending = True Or Receiving = True Then
    MsgBox "you're already sending or receiving", vbOKOnly Or vbCritical, "error"
    Exit Sub
  End If
  
  If PathFileExists(Text5.Text) = 0 Then
    MsgBox "the file you specified does not exist", vbOKOnly Or vbCritical, "error"
    Exit Sub
  End If
  
  Label4.Caption = "loading file into memory..."
  Text5.Enabled = False
  fileSize = FileLen(Text5.Text)
  fileData = String(fileSize, Chr(0))
  Open Text5.Text For Binary As #1
    Get #1, , fileData
  Close #1
  Label4.Caption = "waiting for remote computer to accept file"
  Winsock1.SendData "f=" & Mid(Text5.Text, InStrRev(Text5.Text, "\") + 1)
  Text5.Enabled = True
  
End If

End Sub

Private Sub Command5_Click()
Winsock1.Close
Label4.Caption = "disconnect @ " & Now
Reset
fileSize = 0
Sending = False
Receiving = False
Text1.Enabled = True

End Sub

Private Sub Form_Activate()
Dim c As String
c = Command
If InStr(1, c, "-min") <> 0 Then Form1.WindowState = 1

End Sub

Private Sub Form_Load()



Dim c As String
c = Command
ChDir App.Path

If InStr(1, c, "-noload") <> 0 Then noLoad = True
If InStr(1, c, "-nosave") <> 0 Then noSave = True
If InStr(1, c, "-acceptconnections") <> 0 Then acceptConnections = True
If InStr(1, c, "-accepttransfers") <> 0 Then acceptTransfers = True

If PathFileExists("Odesa.ini") = 1 And Not noLoad Then
  Dim s As String
  Open "Odesa.ini" For Input As #1
    DoEvents
    Line Input #1, s
    Lft = Val(s)
    Line Input #1, s
    Tp = Val(s)
    
    Line Input #1, s
    Text1.Text = s
    Line Input #1, s
    Text2.Text = s
    Line Input #1, s
    Text3.Text = s
    Line Input #1, s
    Text4.Text = s
  Close #1
End If

If InStr(1, c, "-listenport=") <> 0 Then Text2 = Val(Mid(c, InStr(1, c, "-listenport=") + 12))
If InStr(1, c, "-savedir=") <> 0 Then Text1 = Mid(c, InStr(1, c, "-savedir=") + 10, InStr(InStr(1, c, "-savedir=") + 10, c, "'") - (InStr(1, c, "-savedir=") + 10))
If InStr(1, c, "-listen") <> 0 Then autoListen = True
If autoListen Then Command1_Click


End Sub

Private Sub Form_Resize()
If Lft = -5 And Tp = -5 Then Exit Sub
Form1.Left = Lft
Form1.Top = Tp
Lft = -5
Tp = -5

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not noSave Then
  ChDir App.Path
  Open "Odesa.ini" For Output As #1
    Print #1, CStr(Form1.Left)
    Print #1, CStr(Form1.Top)
    
    Print #1, Text1.Text
    Print #1, Text2.Text
    Print #1, Text3.Text
    Print #1, Text4.Text
  Close #1
End If

End Sub

Private Sub Winsock1_Close()
If Left(Label4.Caption, 12) <> "disconnected" Then Label4.Caption = "disconnect @ " & Now
Reset
Beep
Sending = False
Receiving = False
Text1.Enabled = True
If autoListen Then Command1_Click

End Sub

Private Sub Winsock1_Connect()
Label4.Caption = "waiting for remote computer to accept..."

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID

Beep
If acceptConnections Then GoTo skip0
If MsgBox("accept connection from " & Winsock1.RemoteHostIP & "?", vbYesNo Or vbQuestion, "accept?") = vbNo Then
  Winsock1.Close
  Label4.Caption = "rejected remote computer"
Else
skip0:
  If PathIsDirectory(Text1.Text) = 0 Then
    MsgBox "before you can listen or connect, your save directory must be valid", vbOKOnly Or vbCritical, "error"
    Label4.Caption = "invalid save directory"
    Text1.SetFocus
    Winsock1.Close
    Exit Sub
  End If
  If Right(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"
  
  Label4.Caption = "connect @ " & Now
  Winsock1.SendData "c=" & CStr(App.major) & "." & CStr(App.minor) & " " & CStr(App.revision)
  Text1.Enabled = False
End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dR As String, a As VbMsgBoxResult
Winsock1.GetData dR

If Sending = True Then
  Winsock1.SendData fileData
  DoEvents
  Exit Sub
End If

If Receiving = True Then
  Put #1, iWritePos, dR
  iWritePos = iWritePos + Len(dR)
  Label4.Caption = "recieving file... " & Format((fileSize - iWritePos) / 1000000, "0.0") & " Mbs remaining... " & CStr(Int(100 * (iWritePos / fileSize))) & "%"
  DoEvents
  If iWritePos >= fileSize Then
    Close #1
    Label4.Caption = "file recieved"
    Receiving = False
    fileSize = 0
  End If
  Exit Sub
End If

If fileSize = 0 Then
  If Left(dR, 2) = "s=" Then
    fileSize = Val(Mid(dR, 3))
    fileData = ""
    Receiving = True
    Winsock1.SendData "send"
    
  ElseIf Left(dR, 2) = "f=" Then
    Beep
    If acceptTransfers Then GoTo skip0
    a = MsgBox("accept file: " & Mid(dR, 3), vbYesNo Or vbQuestion, "accept?")
    If a = vbYes Then
skip0:
      lFileName = Text1.Text & Mid(dR, 3)
      Open lFileName For Binary As #1
      iWritePos = 1
      Label4.Caption = "recieving file... 0%"
      Winsock1.SendData "sendSize"
      
    Else
      Winsock1.SendData "no"
      Label4.Caption = "denied file"
    End If
    
  ElseIf Left(dR, 2) = "c=" Then
    Dim v As Version
    v.major = Mid(dR, 3, InStr(1, dR, ".") - 3)
    v.minor = Mid(dR, InStr(1, dR, ".") + 1, InStr(1, dR, " ") - InStr(1, dR, ".") - 1)
    v.revision = Mid(dR, InStr(1, dR, " ") + 1)
    
    If App.major <> v.major Or App.minor <> v.minor Or App.revision <> v.revision Then
      Winsock1.SendData "dsc_ver"
      Label4.Caption = "disconnected... versions incompatible"
      Exit Sub
    End If
    
    Label4.Caption = "connect @ " & Now
    
  ElseIf dR = "dsc_ver" Then
    Winsock1.Close
    Label4.Caption = "disconnected... versions incompatible"
    Exit Sub
  End If
  
Else
  If dR = "sendSize" Then
    Winsock1.SendData "s=" & fileSize
    Label4.Caption = "negotiating transfer..."
    
  ElseIf dR = "no" Then
    Label4.Caption = "remote computer denied the file"
    fileSize = 0
    
  ElseIf dR = "send" Then
    Sending = True
    sProgress = 0
    Winsock1.SendData fileData
    Label4.Caption = "sending... 0%"
  End If
End If

End Sub

Private Sub Winsock1_SendComplete()
If Sending = False Then Exit Sub
Sending = False
fileSize = 0
Label4.Caption = "transfer complete"

End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If Sending = False Then Exit Sub

sProgress = sProgress + bytesSent
Label4.Caption = "sending... " & Format(bytesRemaining / 1000000, "0.0") & " Mbs remaining... " & CStr(Int(100 * (sProgress / (sProgress + bytesRemaining)))) & "%"
DoEvents

End Sub


