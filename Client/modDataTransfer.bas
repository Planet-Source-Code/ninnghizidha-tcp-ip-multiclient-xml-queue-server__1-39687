Attribute VB_Name = "modXMLData"
Option Explicit


Public Function ParseData(ByVal xmlstr As String) As String
    
    Dim doc As DOMDocument
    Dim nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim xmlattribute As IXMLDOMAttribute
    Dim strName As String
    Dim strData As String
    Dim strTempReturnValue
    
    Set doc = New DOMDocument
    doc.loadXML xmlstr   'load xml string
    
    'output all Messages from the Clients:
    Set nodelist = doc.getElementsByTagName("clientmessage")
    'MsgBox nodelist.length
    For Each node In nodelist
        For Each xmlattribute In node.Attributes
                Select Case xmlattribute.Name
                    Case "name":   strName = xmlattribute.Value
                    Case "data":   strData = xmlattribute.Value
                End Select
       Next xmlattribute
       strTempReturnValue = strTempReturnValue & strName & ": " & strData & vbCrLf
    Next node
    
    ' and now the Serverbroadcasts
    Set nodelist = doc.getElementsByTagName("broadcast")
    'MsgBox nodelist.length
    For Each node In nodelist
        For Each xmlattribute In node.Attributes
                Select Case xmlattribute.Name
                    Case "name":   strName = xmlattribute.Value
                    Case "data":   strData = xmlattribute.Value
                End Select
       Next xmlattribute
       strTempReturnValue = strTempReturnValue & strName & ": " & strData & vbCrLf
    Next node
    
    Set doc = Nothing
    Set nodelist = Nothing
    Set node = Nothing
    ParseData = strTempReturnValue
    
End Function



' This would be fine for an Module.
Public Sub SendData(pstrSendData As String)
Dim doc As DOMDocument
Dim node As IXMLDOMNode, declPI As IXMLDOMNode
Dim namenode As IXMLDOMNode
    
    
    Set doc = New DOMDocument 'initializing DOMDocument
    
    Set declPI = doc.createProcessingInstruction("xml", " version=""1.0"" ")
    doc.appendChild declPI
    
    
    Set node = doc.createNode(NODE_ELEMENT, "client", "") 'create node Friend
    doc.appendChild node
    
    Set namenode = doc.createNode(NODE_ATTRIBUTE, "name", "")
    namenode.Text = frmMain.hWnd
    doc.childNodes(1).Attributes.setNamedItem namenode
    
    Set namenode = doc.createNode(NODE_ATTRIBUTE, "data", "")
    namenode.Text = pstrSendData
    doc.childNodes(1).Attributes.setNamedItem namenode
    
    frmMain.wsClient.SendData doc.xml
    
    
    Set doc = Nothing

End Sub

