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
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3360
   End
   Begin VB.TextBox TextResult 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0006
      Top             =   1560
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
    'Dim postData As String
    'Dim url As String
    'Dim HttpClient As Object

    'url = "http://192.168.8.22:6394/ws/r/aws_ttsrv2?wsdl"
    'postData = TextInput.Text
    'postData = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tip=""http://www.dsc.com.tw/tiptop/TIPTOPServiceGateWay"">" & _
    '            "<soapenv:Header/><soapenv:Body><tip:GetCsfi020Request><tip:request>" & _
    '            "&lt;Request>&lt;Access>&lt;Authentication user=""tiptop"" password=""tiptop"" />&lt;Connection application="""" source=""192.168.8.22"" />&lt;Organization name=""echom_gz"" />&lt;Locale language=""zh_cn"" />&lt;/Access>" & _
    '            "&lt;RequestContent>&lt;Document>&lt;RecordSet id=""1"" >&lt;Master name=""tc_sfh_file"">&lt;Record>&lt;Field name=""tc_sfh04"" value=""7930B65P500746MMZA500001""/>&lt;/Record>&lt;/Master>&lt;/RecordSet>&lt;/Document>&lt;/RequestContent>" & _
    '            "&lt;/Request></tip:request></tip:GetCsfi020Request></soapenv:Body></soapenv:Envelope>"

    'Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    'HttpClient.open "POST", url, False
    'HttpClient.setRequestHeader "Content-Type", "application/xml;charset=UTF-8"
    'HttpClient.setRequestHeader "SOAPAction", """"
    'HttpClient.send pvToByteArray(postData)
    
    'Do While HttpClient.readyState <> 4
    '    DoEvents
    'Loop
    
    'TextResult.Text = HttpClient.Status
    Dim objHTTP As New MSXML.XMLHTTPRequest
    Dim strEnvelope As String
    Dim strReturn As String
    Dim objReturn As New MSXML.DOMDocument
    Dim dblTax As Double
    Dim strQuery As String
    
    strEnvelope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tip=""http://www.dsc.com.tw/tiptop/TIPTOPServiceGateWay"">" & _
                "<soapenv:Header/><soapenv:Body><tip:GetCsfi020Request><tip:request>" & _
                "&lt;Request>&lt;Access>&lt;Authentication user=""tiptop"" password=""tiptop"" />&lt;Connection application="""" source=""192.168.8.22"" />&lt;Organization name=""echom_gz"" />&lt;Locale language=""zh_cn"" />&lt;/Access>" & _
                "&lt;RequestContent>&lt;Document>&lt;RecordSet id=""1"" >&lt;Master name=""tc_sfh_file"">&lt;Record>&lt;Field name=""tc_sfh04"" value=""7930B65P500746MMZA500001""/>&lt;/Record>&lt;/Master>&lt;/RecordSet>&lt;/Document>&lt;/RequestContent>" & _
                "&lt;/Request></tip:request></tip:GetCsfi020Request></soapenv:Body></soapenv:Envelope>"

    'Set up to post to our local server
    objHTTP.open "POST", "http://192.168.8.22:6394/ws/r/aws_ttsrv2?wsdl", False

    'Set a standard SOAP/ XML header for the content-type
    objHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
    objHTTP.setRequestHeader "Content-Length", Len(strEnvelope)
    TextInput.Text = Len(strEnvelope)
    objHTTP.setRequestHeader "SOAPAction", """"""

    'Set a header for the method to be called
    'objHTTP.setRequestHeader "SOAPMethodName", "urn:tip#GetCsfi020Request"

    'Make the SOAP call
    objHTTP.send strEnvelope

    'Get the return envelope
    'strReturn = objHTTP.responseText
    TextResult.Text = objHTTP.Status & ": " & objHTTP.statusText

    'Load the return envelope into a DOM
    'objReturn.loadXML strReturn

    'Query the return envelope
    'strQuery = _
    '    "SOAP:Envelope/SOAP:Body/m:GetSalesTaxResponse/SalesTax"
    '    dblTax = objReturn.selectSingleNode(strQuery).Text
    MsgBox "End"
End Sub

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = GB2312ToUTF8(sText)
End Function

Public Function GB2312ToUTF8(strIn As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Variant
    Dim adoStream As Object
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 2 'adTypeText
    adoStream.open
    adoStream.WriteText strIn
    adoStream.Position = 0
    adoStream.Type = 1 'adTypeBinary
    GB2312ToUTF8 = adoStream.Read()
    adoStream.Close
    
    If ReturnValueType = vbString Then GB2312ToUTF8 = Mid(GB2312ToUTF8, 1)
End Function
