VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckConnect1730 
      BackColor       =   &H00E0E0E0&
      Caption         =   "连接IO卡"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   2880
      Picture         =   "frmSplash.frx":1DF72
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   4
      Top             =   120
      Width           =   780
   End
   Begin VB.PictureBox PictureBrand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   360
      Picture         =   "frmSplash.frx":1FD66
      ScaleHeight     =   750
      ScaleWidth      =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2550
   End
   Begin VB.ComboBox cmbModelName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1440
      Width           =   3300
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "请选择机型:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrExit

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    sqlstring = "select * from CheckItem"
    Executesql (sqlstring)

    If rs.EOF = False Then
        rs.MoveFirst
        cmbModelName.Clear

        Do While Not rs.EOF
            cmbModelName.AddItem rs.Fields("Mark")
            rs.MoveNext
        Loop
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
        End
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)
    
    If rs.EOF = False Then
        strCurrentModelName = rs("CurrentModelName")
        SetTVCurrentComBaud = rs("ComBaud")
        SetTVCurrentComID = rs("ComID")
        IsStepTime = rs("Delayms")
        delayMs01 = rs("DelayMs01")
        delayMs02 = rs("DelayMs02")
        port1730 = rs("1730Port")
        strErpUrl = rs("ErpUrl")
        strErpOrganization = rs("ErpOrganization")

        If rs("CommunicationMode") = "UART" Then
            isUartMode = True
        Else
            isUartMode = False
        End If
        If rs("Connect1730") Then
            CheckConnect1730.Value = 1
        Else
            CheckConnect1730.Value = 0
        End If
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    cmbModelName.Text = strCurrentModelName
    strCurrentModelName = cmbModelName.Text

    Exit Sub
    
ErrExit:
    MsgBox err.Description, vbCritical, err.Source
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit
    Dim i As Integer

    strCurrentModelName = cmbModelName.Text
    sqlstring = ""
    
    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)
    
    barcodeLen = rs("SN_Len")
    
    For i = 0 To itemNumOfTvInfo
        chkTitleFlag(i) = rs.Fields(i + 14)
    Next i
    
    For i = 0 To 11
        strTvInfoSpec(i) = rs.Fields(i + 2)
    Next i

    rs.Update

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)
    
    If CheckConnect1730.Value = 1 Then
        rs.Fields(6) = True
        isConnect1730 = True
    ElseIf CheckConnect1730.Value = 0 Then
        rs.Fields(6) = False
        isConnect1730 = False
    End If
    
    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""
    
    Form1.Show

    Exit Sub
    
ErrExit:
    MsgBox err.Description, vbCritical, err.Source
End Sub
