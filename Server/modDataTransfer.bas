Attribute VB_Name = "modXMLData"
Option Explicit

Public gdomXMLdoc         As DOMDocument  'This one will be filled with the cleintmessages
Public gdomXMLdocTemplate As DOMDocument  'This is the clean template


Public Sub ParseXMLData(ByVal xmlstr As String)
    
    Dim doc As DOMDocument
    Dim nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim xmlattribute As IXMLDOMAttribute
    Dim strName As String, strData As String
    
    Set doc = New DOMDocument
    doc.loadXML xmlstr   'load xml string
    
    'Just use the Clients
    Set nodelist = doc.getElementsByTagName("client")
    'MsgBox nodelist.length
    For Each node In nodelist
        
        'look at the XML-Attributes of the client
        For Each xmlattribute In node.Attributes
                Select Case xmlattribute.Name
                    Case "name":   strName = xmlattribute.Value
                    Case "data":   strData = xmlattribute.Value
                End Select
       Next xmlattribute
       AddDataToXML "clientmessage", strName, strData 'Add the lientmessage to the queue
    Next node
    
    Set doc = Nothing
    Set nodelist = Nothing
    Set node = Nothing
End Sub


Public Function AddDataToXML(strNode As String, strName As String, strData As String)
    
    Dim node As IXMLDOMNode
    Dim xmlattribute As IXMLDOMAttribute
    Dim namenode As IXMLDOMNode
    
    
    Set node = gdomXMLdoc.createNode(NODE_ELEMENT, strNode, "")
    gdomXMLdoc.childNodes(1).appendChild node

    
    Set xmlattribute = gdomXMLdoc.createAttribute("name")
    xmlattribute.nodeValue = strName
    
    gdomXMLdoc.childNodes(1).childNodes(gdomXMLdoc.childNodes(1).childNodes.length - 1).Attributes.setNamedItem xmlattribute
    
    Set xmlattribute = gdomXMLdoc.createAttribute("data")
    xmlattribute.nodeValue = strData
    gdomXMLdoc.childNodes(1).childNodes(gdomXMLdoc.childNodes(1).childNodes.length - 1).Attributes.setNamedItem xmlattribute
    
    
End Function



Public Function ServerSendData(pstrSendData As String)
' cycle thru all clients, sending the same data.
Dim i As Integer

    For i = 0 To frmMain.wsClients.UBound
        If frmMain.wsClients(i).State = 7 Then
            frmMain.wsClients(i).SendData pstrSendData
            DoEvents ' important: * Not to lock the PC
                     '            * There seems to been a bug with winsock.
        End If
    Next i
End Function
