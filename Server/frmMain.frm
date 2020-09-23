VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "TCP/IP Multiuser XML Queue-Server"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   465
      Left            =   150
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   4290
      TabIndex        =   8
      Top             =   765
      Width           =   4290
   End
   Begin VB.TextBox txtSent 
      Height          =   2055
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtReceived 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox txtErrors 
      Height          =   975
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock wsClients 
      Index           =   0
      Left            =   8400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9001
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   8040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2400
   End
   Begin VB.Timer tmrFlooding 
      Interval        =   15000
      Left            =   7200
      Top             =   0
   End
   Begin VB.Timer tmrConnected 
      Interval        =   500
      Left            =   6840
      Top             =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Outbound (1 XML all 15 seconds)"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Send Broadcast"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      Index           =   0
      X1              =   120
      X2              =   9000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      Caption         =   "Inbound (lots of XML)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Error Log"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim node As IXMLDOMNode, declPI As IXMLDOMNode
Dim namenode As IXMLDOMNode
    
    
    wsServer.LocalPort = 9000
    wsServer.RemotePort = 0
    wsServer.Listen
    ' We Listen to Port 9000, and we don't need a remotehost for it.
    
    
    Set gdomXMLdocTemplate = New DOMDocument 'initializing DOMDocument
    Set gdomXMLdoc = New DOMDocument
    
    Set declPI = gdomXMLdocTemplate.createProcessingInstruction("xml", " version=""1.0"" ")
    gdomXMLdocTemplate.appendChild declPI
    
    
    Set node = gdomXMLdocTemplate.createNode(NODE_ELEMENT, "servermessage", "") 'create node Friend
    gdomXMLdocTemplate.appendChild node
    
    Set namenode = gdomXMLdocTemplate.createNode(NODE_ATTRIBUTE, "ip", "")
    namenode.Text = wsServer.LocalIP
    gdomXMLdocTemplate.childNodes(1).Attributes.setNamedItem namenode
    
    Set namenode = gdomXMLdocTemplate.createNode(NODE_ATTRIBUTE, "name", "")
    namenode.Text = wsServer.LocalHostName
    gdomXMLdocTemplate.childNodes(1).Attributes.setNamedItem namenode
    
    Set namenode = gdomXMLdocTemplate.createNode(NODE_ATTRIBUTE, "time", "")
    namenode.Text = Now()
    gdomXMLdocTemplate.childNodes(1).Attributes.setNamedItem namenode
    
    
    gdomXMLdoc.loadXML gdomXMLdocTemplate.xml
    
End Sub



Private Sub tmrConnected_Timer()
Dim i As Integer, intCon As Integer

    intCon = 0
    ' cycle thru all clients
    For i = 0 To wsClients.UBound
        If wsClients(i).State = 7 Then ' if conencted
            intCon = intCon + 1        ' count the client
        End If
    Next i
        
    Me.Caption = "TCP/IP Multiuser XML Queue-Server - " & intCon & " clients serving"
    ' Just show, how many Cons are at this time
 
    
End Sub

Private Sub tmrFlooding_Timer()
    ServerSendData gdomXMLdoc.xml
    txtSent.SelStart = Len(txtSent.Text)
    txtSent.SelText = gdomXMLdoc.xml & vbCrLf
    gdomXMLdoc.loadXML gdomXMLdocTemplate.xml 'Copy it back, so we have a clean XML-file
End Sub

Private Sub txtMessage_KeyPress(KeyCode As Integer)
Dim i As Integer
    
    If KeyCode = 13 And Trim(txtMessage.Text) <> "" Then ' enter and not Empty
        AddDataToXML "broadcast", "Server", txtMessage.Text
        ' direct broadcast
        txtMessage.Text = ""
        ' no print on the Server-Chatbox - Server don't need to see
        ' what it sended ^^
        KeyCode = 0
        ' Don't pleep, please
    End If

End Sub

Private Sub wsClients_DataArrival(index As Integer, ByVal bytesTotal As Long)
' called, if Server gets data from the clients.
Dim Data As String

    wsClients(index).GetData Data
    
    
    
    
    ' get the Clientdata and put it in the queue
    ParseXMLData Data

     
    txtReceived.SelStart = Len(txtReceived.Text)
    txtReceived.SelText = Data & vbCrLf
    ' Print it.

End Sub


Private Sub wsServer_ConnectionRequest(ByVal requestID As Long)
' if an Client requests a connection
Dim index As Integer
    
    index = GetOpenWinsock
    wsClients(index).Accept requestID
    ' attention:
    ' Request at wsServer, accept at wsClient
End Sub



Private Function GetOpenWinsock() As Integer
'Searches the first open Winsock for us.
Static intUsedPorts As Integer
Dim i As Integer, bOpenWinSockFound As Boolean

bOpenWinSockFound = False
    
    For i = wsClients.UBound To 0 Step -1
        If wsClients(i).State = 0 And Not bOpenWinSockFound Then
            bOpenWinSockFound = True
            GetOpenWinsock = i
        End If
    Next i

    If Not bOpenWinSockFound Then
    ' no open ClientSock found.
        Load wsClients(wsClients.UBound + 1)
        ' load a new Client-winsock into tha array
        intUsedPorts = intUsedPorts + 1 ' new port
        wsClients(wsClients.UBound).LocalPort = wsClients(wsClients.UBound).LocalPort + intUsedPorts

        GetOpenWinsock = wsClients.UBound
        'return the new winsock in the array.
    End If


End Function



Private Sub wsClients_Error(index As Integer, ByVal iNumber As Integer, strDescription As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'called, if an error occured
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsClients(" & index & ") - " & iNumber & " - " & strDescription & vbCrLf
    wsClients(index).Close
    ' Close the Winsock on which the error happend.
End Sub
Private Sub wsServer_Error(ByVal iNumber As Integer, strDescription As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' same as above for wsServer-Winsock
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsServer - " & iNumber & " - " & strDescription & vbCrLf
End Sub







