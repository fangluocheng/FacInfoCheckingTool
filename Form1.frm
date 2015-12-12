VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "工厂信息校验工具"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5790
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "退出老化"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "进入老化"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   0
      Text            =   "123456789"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sampl1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu vbSet 
      Caption         =   "设置"
      Begin VB.Menu tbSetComPort 
         Caption         =   "设置串口"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "设置条码长度"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RES As Long
Dim Result As Boolean
Dim StepTime As Long

Private Sub Command2_Click()
    Call SET_BURNING_MODE(1)
End Sub

Private Sub Command3_Click()
    Call SET_BURNING_MODE(0)
    EXIT_FAC_MODE
End Sub

Private Sub Form_Load()

    i = 0

    SetTVCurrentComBaud = 115200

    StepTime = IsStepTime
    IsStop = False
    subInitComPort
    subInitInterface

    'Label8 = strCurrentModelName

End Sub

Private Sub subInitInterface()
    txtInput.Text = ""
End Sub

Private Sub subInitComPort()
    'sqlstring = "select * from CommonTable where Mark='ATS'"
    'Executesql (sqlstring)

    'If rs.EOF = False Then
    '    SetTVCurrentComID = rs("ComID")
    'Else
    '    MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    'End
    'End If

    'Set cn = Nothing
    'Set rs = Nothing
    'sqlstring = ""

    SetTVCurrentComID = 6
    ComInit
    
    ENTER_FAC_MODE

End Sub


Private Sub tbSetComPort_Click()
    Form2.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    IsStop = False
    'subMainProcesser
     
    If IsStop = True Then
        Exit Sub
    End If
End Sub

Private Sub vbSetSPEC_Click()
    frmSetData.Show
End Sub
