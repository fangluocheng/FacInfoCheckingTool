VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "参数设置"
   ClientHeight    =   6630
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetData.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11070
   Begin VB.Frame Frame3 
      Caption         =   "通讯模式"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7080
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
      Begin VB.OptionButton optNetwork 
         Caption         =   "网络"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1200
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optUart 
         Caption         =   "串口"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TV 信息"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4700
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   10820
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   14
         Left            =   7185
         TabIndex        =   42
         Text            =   "----"
         Top             =   4140
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   13
         Left            =   3645
         TabIndex        =   41
         Text            =   "None"
         Top             =   4140
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Text            =   "None"
         Top             =   4140
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "分区版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   12
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "区域"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   13
         Left            =   3645
         TabIndex        =   38
         Top             =   3720
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "Device Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   14
         Left            =   7185
         TabIndex        =   37
         Top             =   3720
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   11
         Left            =   7185
         TabIndex        =   36
         Text            =   "----"
         Top             =   3270
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   10
         Left            =   3645
         TabIndex        =   35
         Text            =   "None"
         Top             =   3270
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Text            =   "----"
         Top             =   3270
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "HDCP Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   2850
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "4K/2K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   10
         Left            =   3645
         TabIndex        =   32
         Top             =   2850
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "MAC 地址"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   11
         Left            =   7185
         TabIndex        =   31
         Top             =   2850
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   8
         Left            =   7185
         TabIndex        =   30
         Text            =   "None"
         Top             =   2400
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   7
         Left            =   3645
         TabIndex        =   29
         Text            =   "None"
         Top             =   2400
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Text            =   "None"
         Top             =   2400
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "2.4G 版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   1980
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "屏型号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   7
         Left            =   3645
         TabIndex        =   26
         Top             =   1980
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "播控平台"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   8
         Left            =   7185
         TabIndex        =   25
         Top             =   1980
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   7185
         TabIndex        =   24
         Text            =   "None"
         Top             =   1530
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   3645
         TabIndex        =   23
         Text            =   "None"
         Top             =   1530
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Text            =   "None"
         Top             =   1530
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "硬件版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1110
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   4
         Left            =   3645
         TabIndex        =   20
         Top             =   1110
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "渠道"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   5
         Left            =   7185
         TabIndex        =   19
         Top             =   1110
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "Flash 信息"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   2
         Left            =   7180
         TabIndex        =   14
         Top             =   240
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "系统版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   1
         Left            =   3650
         TabIndex        =   13
         Top             =   240
         Width           =   3500
      End
      Begin VB.CheckBox chkTitle 
         BackColor       =   &H00808080&
         Caption         =   "机型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   410
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Text            =   "None"
         Top             =   660
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   3650
         TabIndex        =   10
         Text            =   "None"
         Top             =   660
         Width           =   3500
      End
      Begin VB.TextBox txtTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   7180
         TabIndex        =   9
         Text            =   "None"
         Top             =   660
         Width           =   3500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "串口设置"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   6855
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1000
         TabIndex        =   0
         Text            =   "115200"
         Top             =   300
         Width           =   950
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3610
         TabIndex        =   1
         Text            =   "500"
         Top             =   300
         Width           =   950
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5720
         TabIndex        =   2
         Text            =   "1"
         Top             =   300
         Width           =   950
      End
      Begin VB.Label Label2 
         Caption         =   "波特率"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "延迟时间 (ms)"
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
         Left            =   2200
         TabIndex        =   6
         Top             =   360
         Width           =   1400
      End
      Begin VB.Label Label4 
         Caption         =   "条码长度"
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
         Left            =   4810
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9840
      TabIndex        =   3
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    Dim i As Integer
    
    i = 0

    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    Label1 = strCurrentModelName
    Text1.Text = rs("ComBaud")
    Text2.Text = rs("Delayms")
    Text3.Text = rs("SN_Len")
    
    'Read the Spec data from database and show them into the TextBox
    For i = 0 To (itemNumOfTvInfo - 3)
        If i <= 8 Then
            txtTVInfo(i).Text = rs.Fields(i + 4)
        ElseIf i = 9 Then
            txtTVInfo(10).Text = rs.Fields(i + 4)
        ElseIf i > 9 Then
            txtTVInfo(i + 2).Text = rs.Fields(i + 4)
        End If
    Next i

    'Whether the CheckBox selected or not.
    For i = 0 To itemNumOfTvInfo
        If rs.Fields(i + 16) Then
            chkTitle(i).Value = 1
        Else
            chkTitle(i).Value = 0
        End If
    Next i

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

End Sub

Private Sub cmdSave_Click()
    Dim i As Integer

    i = 0

    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    'Set the text into the CheckItem table of database file.
    rs.Fields(1) = Val(Text1.Text)                         'ComBaud
    rs.Fields(2) = Val(Text2.Text)                         'Delayms
    rs.Fields(3) = Val(Text3.Text)                         'SN_Len

    For i = 0 To itemNumOfTvInfo
        If chkTitle(i).Value = 1 Then
            rs.Fields(i + 16) = True
            
            If Not (i = 9 Or i = 11 Or i = 14) Then
                If i = 10 Then
                    rs.Fields(13) = txtTVInfo(i).Text
                ElseIf i = 12 Or i = 13 Then
                    rs.Fields(i + 2) = txtTVInfo(i).Text
                Else
                    rs.Fields(i + 4) = txtTVInfo(i).Text
                End If
            End If
        ElseIf chkTitle(i).Value = 0 Then
            rs.Fields(i + 16) = False
        End If
    Next i
 
    rs.Update

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    If rs.EOF = False Then
        If optUart.Value = True Then
            rs.Fields(3) = "UART"
        ElseIf optNetwork.Value = True Then
            rs.Fields(3) = "Network"
        Else
            rs.Fields(3) = "UART"
        End If
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End If

    rs.Update
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    MsgBox "Save success!", vbOKOnly, "warning"
    Unload Me
    Unload Form1

End Sub

