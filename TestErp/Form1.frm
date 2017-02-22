VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextInput 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3360
   End
   Begin VB.TextBox TextResult 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取MAC"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim objHTTP As New XMLHTTP
    Dim strEnvelope As String
    Dim strReturn As String
    Dim objReturn As New DOMDocument
    Dim objNodeList As MSXML2.IXMLDOMNodeList
    Dim strStatus As String
    Dim strMacAddr As String
    Dim strActicode As String
    
    strEnvelope = TestWebPost(TextInput.Text)

    'Set up to post to our local server
    objHTTP.open "POST", "http://192.168.8.22:6394/ws/r/aws_ttsrv2?wsdl", False

    'Set a standard SOAP/ XML header
    objHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
    objHTTP.setRequestHeader "SOAPAction", """"""

    'Make the SOAP call
    objHTTP.send strEnvelope

    'Get the return envelope
    strReturn = Replace(Replace(objHTTP.responseText, "&lt;", "<"), "&gt;", ">")
    TextResult.Text = strReturn

    'Load the return envelope into a DOM
    objReturn.loadXML strReturn
    
    'Query the return envelope, then get acticode and MAC Address
    strQuery = "/SOAP-ENV:Envelope/SOAP-ENV:Body/fjs1:GetCsfi020Response/fjs1:response/" & _
                "Response/ResponseContent/Document/RecordSet/Master/Record/Field"
    Set objNodeList = objReturn.selectNodes(strQuery)

    If Not objNodeList Is Nothing Then
        Dim objNode As MSXML2.IXMLDOMNode
            
        For Each objNode In objNodeList
                If Trim(objNode.selectSingleNode("@name").Text) = "status" Then
                    strStatus = Trim(objNode.selectSingleNode("@value").Text)
                End If
                If Trim(objNode.selectSingleNode("@name").Text) = "maccode" Then
                    strMacAddr = Trim(objNode.selectSingleNode("@value").Text)
                End If
                If Trim(objNode.selectSingleNode("@name").Text) = "acticode" Then
                    strActicode = Trim(objNode.selectSingleNode("@value").Text)
                End If
            Next objNode
        End If

    If strStatus = "Y" Then
        If strActicode = "Y" Then
            MsgBox "MAC Address: " & strMacAddr
        ElseIf strActicode = "N" Then
            MsgBox "此整机码无效！"
        End If
    ElseIf strStatus = "N" Then
        MsgBox "在 ERP 系统上找不到这台电视的整机码！"
    End If
End Sub

Public Function createHeaderXML() As String
    createHeaderXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>" & _
                "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tip=""http://www.dsc.com.tw/tiptop/TIPTOPServiceGateWay"">" & _
                "<soapenv:Header/><soapenv:Body><tip:GetCsfi020Request><tip:request>"
End Function

Public Function createEndXML()
    createEndXML = "</tip:request></tip:GetCsfi020Request></soapenv:Body></soapenv:Envelope>"
End Function

Public Function createPartXML(snCode As String) As String
    createPartXML = "&lt;Request&gt; &lt;Access&gt; &lt;Authentication user=""tiptop"" password=""tiptop""/&gt; &lt;Connection application="""" source=""192.168.8.22""/&gt; &lt;Organization name=""echom_gz""/&gt; &lt;Locale language=""zh_cn""/&gt; &lt;/Access&gt;" & _
                "&lt;RequestContent&gt; &lt;Document&gt; &lt;RecordSet id=""1""&gt; &lt;Master name=""tc_sfh_file""&gt; &lt;Record&gt; &lt;Field name=""tc_sfh04"" value=" & _
                """" & snCode & """" & _
                "/&gt; &lt;/Record&gt; &lt;/Master&gt; &lt;/RecordSet&gt; &lt;/Document&gt; &lt;/RequestContent&gt;" & _
                "&lt;/Request&gt;"
End Function

Public Function TestWebPost(snCode As String) As String
    Dim testString As String

    testString = createHeaderXML() + createPartXML(snCode) + createEndXML()

    TestWebPost = testString
End Function
