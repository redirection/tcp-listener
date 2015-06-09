VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form tcpform 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tcp protocol"
   ClientHeight    =   2460
   ClientLeft      =   3675
   ClientTop       =   2580
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3930
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   3480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   25
      TabIndex        =   0
      Top             =   -25
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "reset"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2200
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2050
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   590
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "listen"
         Height          =   255
         Left            =   1520
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Port 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "2000"
         Top             =   240
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         BorderWidth     =   5
         X1              =   0
         X2              =   3840
         Y1              =   50
         Y2              =   50
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3405
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "users:"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "port:"
         Height          =   255
         Left            =   80
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "tcpform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxClient As Integer

Private Sub Command1_Click()
If Command1.Caption = "listen" Then
  Winsock(0).LocalPort = Port.Text
  Winsock(0).Listen
  Command1.Caption = "stop"
  Port.Enabled = False
  Command2.Enabled = True
Else
  Winsock(0).Close
  Command1.Caption = "listen"
  Port.Enabled = True
End If
End Sub

Private Sub Command2_Click()
  If Winsock(0).State > 0 Then Winsock(0).Close
  For i = MaxClient To 1 Step -1
    DoEvents
    If Winsock(i).State > 0 Then
      Winsock(i).Close
      Unload Winsock(i)
    End If
  Next i
  MaxClient = 0
  Command1.Caption = "listen"
  Port.Enabled = True
  Command2.Enabled = False
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn And Len(Text2) > 0 Then
  If Left(Text2, 5) = "/msg " And Len(Text2) > 7 Then
    Temp = Mid(Text2, 6, Len(Text2))
    Where = InStr(Temp, " ")
    User = Mid(Temp, 1, Where - 1)
    Temp = Mid(Temp, Where + 1, Len(Temp))
    If Len(Where) = 0 Or Len(User) = 0 Or Len(Temp) = 0 Or Winsock(User).State <> 7 Then
      Text1.SelText = "*invalid parameters and or user*" & vbCrLf
      Exit Sub
    End If
    Winsock(User).SendData Temp
    Text1.SelText = User & "*" & Winsock(0).LocalHostName & ": " & Temp & vbCrLf
  ElseIf Left(Text2, 6) = "/kick " And Len(Text2) > 6 Then
    Temp = Mid(Text2, 7, Len(Text2))
    Where = InStr(Temp, " ")
    If Where = 0 Then Where = Len(Temp) + 1
    User = Mid(Temp, 1, Where - 1)
    Temp = Mid(Temp, Where + 1, Len(Temp))
    If Len(User) > 0 And Winsock(User).State = 7 Then
      Winsock(User).SendData Temp
      DoEvents
      Winsock(User).Close
      Text1.SelText = "*winsock:" & User & " disconnected*" & vbCrLf
    Else
      Text1.SelText = "*invalid parameters and or user*" & vbCrLf
      Exit Sub
    End If
  Else
    For i = 1 To MaxClient
      DoEvents
      Winsock(i).SendData Text2
    Next i
    Text1.SelText = Winsock(0).LocalHostName & ": " & Text2 & vbCrLf
  End If
  Text1.SelStart = Len(Text1)
  Text2 = ""
End If
End Sub

Private Sub Winsock_Close(Index As Integer)
Winsock(Index).Close
For i = MaxClient To 1 Step -1
  DoEvents
  If Winsock(i).State = 7 Then
    Exit For
  Else
    'clean up loaded winsocks
    Winsock(i).Close
    Unload Winsock(i)
    MaxClient = MaxClient - 1
  End If
Next i
Text1.SelText = Winsock(CurrentClient).RemoteHostIP & ":" & Winsock(CurrentClient).RemotePort & " unloaded:" & CurrentClient & vbCrLf
Label3 = MaxClient

Me.Caption = ""
For i = 1 To MaxClient
  DoEvents
  Me.Caption = Me.Caption & " " & i & ":" & Winsock(i).State
Next i
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim CurrentClient As Integer
For i = MaxClient To 1 Step -1
  DoEvents
  If Winsock(i).State <> 7 Then
    Winsock(i).Close
    CurrentClient = i
  End If
Next i
If CurrentClient = 0 Then
  MaxClient = MaxClient + 1
  CurrentClient = MaxClient
  Load Winsock(MaxClient)
End If
Winsock(CurrentClient).Accept requestID
Text1.SelText = Winsock(CurrentClient).RemoteHostIP & ":" & Winsock(CurrentClient).RemotePort & " loaded:" & CurrentClient & vbCrLf
Label3 = MaxClient

Me.Caption = ""
For i = 1 To MaxClient
  DoEvents
  Me.Caption = Me.Caption & " " & i & ":" & Winsock(i).State
Next i
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim MessageBuffer As String
Winsock(Index).GetData MessageBuffer
Text1.SelText = Index & "[" & Winsock(Index).RemoteHostIP & "] " & MessageBuffer & vbCrLf
Text1.SelStart = Len(Text1)
End Sub
