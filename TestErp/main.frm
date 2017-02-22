VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H80000001&
   Caption         =   "SOAPUSER.COM : Quotation Service"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton quit_button 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   26
      Top             =   7400
      Width           =   975
   End
   Begin VB.CommandButton disclaimer_button 
      Caption         =   "Disclaimer"
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   7400
      Width           =   1095
   End
   Begin VB.Frame qlist_frame 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12855
      Begin VB.Frame gqba_frame 
         Height          =   735
         Left            =   2640
         TabIndex        =   5
         Top             =   3720
         Width           =   9975
         Begin VB.CommandButton getQuotationsByAuthor_button 
            Caption         =   "Get Quotations By Author"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox getAuthor_input 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   6
            Text            =   "Wilde, Oscar"
            Top             =   240
            Width           =   6015
         End
         Begin VB.Label getAuthor_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Author :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   8
            Top             =   285
            Width           =   735
         End
      End
      Begin VB.Frame gaq_frame 
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   3720
         Width           =   2415
         Begin VB.CommandButton getAllQuotations_button 
            Caption         =   "Get All Quotations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.ListBox qlist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   12375
      End
      Begin VB.ListBox hiddenlist 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9000
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label qlist_label 
         Alignment       =   2  'Center
         Caption         =   "Quotation List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   12375
      End
   End
   Begin VB.Frame submit_frame 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   12855
      Begin VB.CommandButton submitQuotation_button 
         Caption         =   "Submit Quotation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox submitAuthor_input 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox submitText_input 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   11535
      End
      Begin VB.Label authorHelpLabel3 
         Caption         =   "examples: Wilde, Oscar  |  Kennedy, John F.  |  Elliot, T. S."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   920
         Width           =   4455
      End
      Begin VB.Label authorHelpLabel2 
         Caption         =   "Surname, First Names and/or Initials"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   920
         Width           =   3015
      End
      Begin VB.Label authorHelpLabel1 
         Caption         =   "For Author enter "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Quotation Submission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   12375
      End
      Begin VB.Label submitText_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Text :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label submitAuthor_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Author :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   525
         Width           =   735
      End
   End
   Begin VB.Frame url_frame 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   12855
      Begin VB.CommandButton help_button 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   21
         Top             =   50
         Width           =   1095
      End
      Begin VB.TextBox url_input 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   11
         Text            =   "http://localhost:8080/soap/servlet/rpcrouter"
         Top             =   50
         Width           =   6615
      End
      Begin VB.Label url_label 
         BackColor       =   &H80000001&
         Caption         =   "Soap Server URL :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   105
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "SOAPUSER.COM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   7320
      Width           =   10335
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ENC = "http://schemas.xmlsoap.org/soap/encoding/"
Private Const XSI = "http://www.w3.org/1999/XMLSchema-instance"
Private Const XSD = "http://www.w3.org/1999/XMLSchema"


Private Sub SetSoapHeader(ByRef Connector As SoapConnector, _
                            ByRef Serializer As SoapSerializer, _
                            ByVal Service As String, _
                            ByVal Method As String)

    URI = "urn:" & Service
    Connector.Property("EndPointURL") = url_input.text
    Call Connector.Connect
    Connector.Property("SoapAction") = URI & "#" & Method
    Call Connector.BeginMessage
    Serializer.Init Connector.InputStream
    Serializer.StartEnvelope , ENC
    Serializer.SoapNamespace "xsi", XSI
    Serializer.SoapNamespace "SOAP-ENC", ENC
    Serializer.SoapNamespace "xsd", XSD
    Serializer.StartBody
    Serializer.StartElement Method, URI, , "method"

End Sub


Private Sub SetSoapFooter(ByRef Connector As SoapConnector, _
                            ByRef Serializer As SoapSerializer)

    Serializer.EndElement
    Serializer.EndBody
    Serializer.EndEnvelope
    Connector.EndMessage
    
End Sub


Private Sub GetTagValueArray(ByRef node As MSXML2.IXMLDOMNode, _
                                ByVal tag As String, _
                                ByRef a() As String)

    Dim childnode As MSXML2.IXMLDOMNode

    If node.baseName = tag Then
        On Error GoTo ErrorHandler
        ReDim Preserve a(UBound(a) + 1)
        a(UBound(a)) = node.text
    End If

    If node.childNodes.length > 0 Then
        For Each childnode In node.childNodes
            GetTagValueArray childnode, tag, a
        Next
    End If

    Exit Sub

ErrorHandler:
    ReDim a(0)
    Resume Next

End Sub


Private Sub getAllQuotations()

    Dim Connector As SoapConnector
    Dim Serializer As SoapSerializer
    Dim Reader As SoapReader

    Set Connector = New HttpConnector
    Set Serializer = New SoapSerializer
    Set Reader = New SoapReader

    Dim texts() As String
    Dim authors() As String
    qlist.Clear
    hiddenlist.Clear
    
    SetSoapHeader Connector, Serializer, "QuotationService", _
                                            "getAllQuotations"
    SetSoapFooter Connector, Serializer
    Reader.Load Connector.OutputStream
    
    If Not Reader.Fault Is Nothing Then
        MsgBox Reader.FaultString.text, vbExclamation
    Else
        GetTagValueArray Reader.Dom, "text", texts
        GetTagValueArray Reader.Dom, "author", authors
        On Error Resume Next 'in case arrays not dimensioned
        For i = 0 To UBound(texts)
            qlist.AddItem texts(i) & " (" & authors(i) & ")"
            hiddenlist.AddItem authors(i)
        Next
    End If

End Sub


Private Sub getQuotationsByAuthor(ByVal author As String)

    Dim Connector As SoapConnector
    Dim Serializer As SoapSerializer
    Dim Reader As SoapReader

    Set Connector = New HttpConnector
    Set Serializer = New SoapSerializer
    Set Reader = New SoapReader

    Dim texts() As String
    Dim authors() As String
    qlist.Clear
    hiddenlist.Clear
    
    SetSoapHeader Connector, Serializer, "QuotationService", _
                                            "getQuotationsByAuthor"
    Serializer.StartElement "Author"
    Serializer.SoapAttribute "type", , "xsd:string", "xsi"
    Serializer.WriteString author
    Serializer.EndElement
    SetSoapFooter Connector, Serializer
    Reader.Load Connector.OutputStream
    
    If Not Reader.Fault Is Nothing Then
        MsgBox Reader.FaultString.text, vbExclamation
    Else
        GetTagValueArray Reader.Dom, "item", texts
        On Error Resume Next 'in case arrays not dimensioned
        For i = 0 To UBound(texts)
            qlist.AddItem texts(i)
            hiddenlist.AddItem author
        Next
    End If

End Sub


Private Sub submitQuotation(ByVal author As String, _
                            ByVal text As String)

    Dim Connector As SoapConnector
    Dim Serializer As SoapSerializer
    Dim Reader As SoapReader

    Set Connector = New HttpConnector
    Set Serializer = New SoapSerializer
    Set Reader = New SoapReader
    
    SetSoapHeader Connector, Serializer, "QuotationService", _
                                            "submitQuotation"
    Serializer.StartElement "Author"
    Serializer.SoapAttribute "type", , "xsd:string", "xsi"
    Serializer.WriteString author
    Serializer.EndElement
    Serializer.StartElement "Text"
    Serializer.SoapAttribute "type", , "xsd:string", "xsi"
    Serializer.WriteString text
    Serializer.EndElement
    SetSoapFooter Connector, Serializer
    Reader.Load Connector.OutputStream
    
    If Not Reader.Fault Is Nothing Then
        MsgBox Reader.FaultString.text, vbExclamation
    Else
        MsgBox "Quotation Submitted"
    End If

End Sub


Private Sub getAllQuotations_button_Click()
    Screen.MousePointer = vbHourglass
    getAllQuotations
    qlist_label = "All Quotations"
    Screen.MousePointer = vbDefault
End Sub


Private Sub getQuotationsByAuthor_button_Click()
    Screen.MousePointer = vbHourglass
    getQuotationsByAuthor getAuthor_input.text
    qlist_label = "Quotations accredited to " & getAuthor_input.text
    Screen.MousePointer = vbDefault
End Sub


Private Sub qlist_Click()

    If qlist.ListIndex >= 0 Then
        getAuthor_input.text = hiddenlist.List(qlist.ListIndex)
    End If

End Sub

Private Sub submitQuotation_button_Click()
    Screen.MousePointer = vbHourglass
    submitQuotation submitAuthor_input.text, submitText_input.text
    Screen.MousePointer = vbDefault
End Sub


Private Sub quit_button_Click()
    
    End

End Sub


Private Sub help_button_Click()

    help.Show

End Sub


Private Sub disclaimer_button_Click()

    disclaimer.Show

End Sub
