VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Client"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrDisconnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   600
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   1800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   9000
   End
   Begin VB.TextBox txtReceived 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Send Message"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Received From Server"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
    
    If Trim(txtRemoteHost.Text) = "" Then
        MsgBox "remotehost empty", vbCritical, "User Error"
    Else
        wsClient.RemoteHost = txtRemoteHost.Text
        wsClient.Connect
        Do Until wsClient.State = 7
            ' 0 = closed, 9 = error
            ' we don't need this Errors.
            If wsClient.State = 0 Or wsClient.State = 9 Then
                MsgBox "Error at connecting!", vbCritical, "Winsock Error"
                'error when we tried to connect
                'jump out.
                Exit Sub
            End If
            DoEvents  'Ugly, but worth it ^^
        Loop
        ' hello Mr. Server.
        SendData "welcome"
        
        ' Disable the connect-Buttons, turn on the timer.
        txtRemoteHost.Enabled = False
        cmdConnect.Enabled = False
        tmrDisconnect.Enabled = True
        
        
    End If
End Sub



Private Sub Form_Load()
    Me.Caption = Me.Caption & " " & Me.hWnd
    ' We dont use nicks, we use the hWnd
    ' saves lots of time when you test it.
End Sub

Private Sub tmrDisconnect_Timer()
'Checks, if disconnected from Server
    If wsClient.State <> 7 Then
        MsgBox ("Server disconnected")
        End
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter-key
       If wsClient.State = 7 Then ' just send when connected

            SendData txtMessage.Text
            txtMessage.Text = ""
            KeyAscii = 0 ' So it don't pleep!
        Else
            MsgBox "Not connected!", vbCritical, "User Error"
            
        End If
    End If
End Sub


Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
' Called, when Data arrives
Dim strData As String
    
    wsClient.GetData strData
    
    txtReceived.SelStart = Len(txtReceived.Text)
    txtReceived.SelText = ParseData(strData)
End Sub

Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' Print the WinSock-Error - use it, if you want to
    'MsgBox "Winsock Error: " & Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub
