VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¹¤³§ÐÅÏ¢Ð£Ñé¹¤¾ß"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12630
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame3 
      Caption         =   "ÌõÂë"
      Height          =   1250
      Left            =   4320
      TabIndex        =   33
      Top             =   5880
      Width           =   8175
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   34
         Text            =   "123456789"
         Top             =   240
         Width           =   7905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "²âÊÔ½á¹û"
      Height          =   1250
      Left            =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   4095
      Begin VB.TextBox txtResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   32
         Text            =   "Checking"
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TV ÐÅÏ¢"
      Height          =   5700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin VB.TextBox txtDeviceKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         TabIndex        =   30
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtMacAddr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         TabIndex        =   28
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtResolution 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         TabIndex        =   25
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtHdcpKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   24
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         TabIndex        =   22
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtCarrier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         TabIndex        =   20
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPanelName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         TabIndex        =   17
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtTwoPointFourVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   16
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPartitionVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   14
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         TabIndex        =   12
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtDimension 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         TabIndex        =   9
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtHWVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   8
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtFlashInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         TabIndex        =   6
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtSysVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         TabIndex        =   4
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtModelInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   2
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Device Key"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   29
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAC µØÖ·"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   27
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HDCP"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4K/2K"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   23
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÇøÓò"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   21
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "²¥¿ØÆ½Ì¨"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   19
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2.4G °æ±¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÆÁÐÍºÅ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   15
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "·ÖÇø°æ±¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÇþµÀ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   11
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ó²¼þ°æ±¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   7
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flash ÐÅÏ¢"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   5
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÏµÍ³°æ±¾"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   3
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "»úÐÍ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4000
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11880
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu vbSet 
      Caption         =   "ÉèÖÃ"
      Begin VB.Menu tbSetComPort 
         Caption         =   "ÉèÖÃ´®¿Ú"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "ÉèÖÃÊý¾Ý¹æ¸ñ"
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


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub MSComm1_OnComm()

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

Private Sub txtResult_Change()

End Sub

Private Sub vbSetSPEC_Click()
    frmSetData.Show
End Sub
