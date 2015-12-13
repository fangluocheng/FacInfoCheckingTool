VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "²ÎÊýÉèÖÃ"
   ClientHeight    =   7830
   ClientLeft      =   6435
   ClientTop       =   3210
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12615
   Begin VB.Frame Frame2 
      Caption         =   "TV ÐÅÏ¢"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   12375
      Begin VB.CheckBox Check15 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   8190
         TabIndex        =   38
         Top             =   4560
         Width           =   4000
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   4150
         TabIndex        =   37
         Top             =   4560
         Width           =   4000
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   120
         TabIndex        =   36
         Top             =   4560
         Width           =   4000
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   8190
         TabIndex        =   35
         Top             =   3480
         Width           =   4000
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   4150
         TabIndex        =   34
         Top             =   3480
         Width           =   4000
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   4000
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   8190
         TabIndex        =   32
         Top             =   2400
         Width           =   4000
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   4150
         TabIndex        =   31
         Top             =   2400
         Width           =   4000
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   4000
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   8190
         TabIndex        =   29
         Top             =   1320
         Width           =   4000
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   4150
         TabIndex        =   28
         Top             =   1320
         Width           =   4000
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   4000
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   8190
         TabIndex        =   26
         Top             =   240
         Width           =   4000
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   4150
         TabIndex        =   25
         Top             =   240
         Width           =   4000
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
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
         Height          =   450
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   4000
      End
      Begin VB.TextBox txtModelInfoSpec 
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
         TabIndex        =   23
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtSysVerSpec 
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
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtFlashInfoSpec 
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
         TabIndex        =   21
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtHWVerSpec 
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
         TabIndex        =   20
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtDimensionSpec 
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
         TabIndex        =   19
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtChannelSpec 
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
         TabIndex        =   18
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtPartitionVerSpec 
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
         TabIndex        =   17
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtTwoPointFourVerSpec 
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
      Begin VB.TextBox txtPanelNameSpec 
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
         TabIndex        =   15
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtCarrierSpec 
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
         TabIndex        =   14
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtAreaSpec 
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
         TabIndex        =   13
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtHdcpKeySpec 
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
         TabIndex        =   12
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtResolutionSpec 
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
         TabIndex        =   11
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtMacAddrSpec 
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
         TabIndex        =   10
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtDeviceKeySpec 
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
         TabIndex        =   9
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "´®¿ÚÉèÖÃ"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   8775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   0
         Text            =   "115200"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Text            =   "500"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "²¨ÌØÂÊ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "ÑÓ³ÙÊ±¼ä"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "ÌõÂë³¤¶È"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "±£´æ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11160
      TabIndex        =   3
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   39
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

    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    Label1 = strCurrentModelName
    Text1.Text = rs("ComBaud")
    Text2.Text = rs("Delayms")
    Text3.Text = rs("SN_Len")
    
    'Read the Spec data from database and show them into the TextBox
    txtModelInfoSpec.Text = rs("ModelSpec")
    txtSysVerSpec.Text = rs("SysVerSpec")
    txtFlashInfoSpec.Text = rs("FlashInfoSpec")
    txtHWVerSpec.Text = rs("HardwareVerSpec")
    txtDimensionSpec.Text = rs("DimensionSpec")
    txtChannelSpec.Text = rs("ChannelSpec")
    txtTwoPointFourVerSpec.Text = rs("TwoPointFourGVerSpec")
    txtPanelNameSpec.Text = rs("PanelSpec")
    txtCarrierSpec.Text = rs("CarrierSpec")
    txtHdcpKeySpec.Text = rs("HDCPSpec")
    txtResolutionSpec.Text = rs("ResolutionSpec")
    txtMacAddrSpec.Text = rs("MACAddrSpec")
    txtPartitionVerSpec.Text = rs("PartitionVerSpec")
    txtAreaSpec.Text = rs("AreaSpec")
    txtDeviceKeySpec.Text = rs("DeviceKeySpec")

    'Whether the CheckBox selected or not.
    If rs("IsModel") Then Check1.Value = 1 Else Check1.Value = 0
    End If
    If rs("IsSysVer") Then Check2.Value = 1 Else Check2.Value = 0
    End If
    If rs("IsFlashInfo") Then Check3.Value = 1 Else Check3.Value = 0
    End If
    If rs("IsHardwareVer") Then Check4.Value = 1 Else Check4.Value = 0
    End If
    If rs("IsDimension") Then Check5.Value = 1 Else Check5.Value = 0
    End If
    If rs("IsChannel") Then Check6.Value = 1 Else Check6.Value = 0
    End If
    If rs("Is24GVer") Then Check7.Value = 1 Else Check7.Value = 0
    End If
    If rs("IsPanel") Then Check8.Value = 1 Else Check8.Value = 0
    End If
    If rs("IsCarrier") Then Check9.Value = 1 Else Check9.Value = 0
    End If
    If rs("IsHDCP") Then Check10.Value = 1 Else Check10.Value = 0
    End If
    If rs("IsResolution") Then Check11.Value = 1 Else Check11.Value = 0
    End If
    If rs("IsMACAddr") Then Check12.Value = 1 Else Check12.Value = 0
    End If
    If rs("IsPartitionVer") Then Check13.Value = 1 Else Check13.Value = 0
    End If
    If rs("IsArea") Then Check14.Value = 1 Else Check14.Value = 0
    End If
    If rs("IsDeviceKey") Then Check15.Value = 1 Else Check15.Value = 0
    End If

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

End Sub

Private Sub Command1_Click()

    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)
  
    rs.Fields(1) = Val(Text1.Text)                         'ComBaud
    rs.Fields(2) = Val(Text2.Text)                         'Delayms
    rs.Fields(3) = Val(Text3.Text)                         'SN_Len

    rs.Fields(4) = Val(txtModelInfoSpec.Text)              'ModelM
    rs.Fields(5) = Val(txtSysVerSpec.Text)                 'SysVerM
    rs.Fields(6) = Val(txtFlashInfoSpec.Text)              'FlashInfoM
    rs.Fields(7) = Val(txtHWVerSpec.Text)                  'HardwareVerM
    rs.Fields(8) = Val(txtDimensionSpec.Text)              'DimensionM
    rs.Fields(9) = Val(txtChannelSpec.Text)                'ChannelM
    rs.Fields(10) = Val(txtTwoPointFourVerSpec.Text)       '24GVerM
    rs.Fields(11) = Val(txtPanelNameSpec.Text)             'PanelM
    rs.Fields(12) = Val(txtCarrierSpec.Text)               'CarrierM
    rs.Fields(13) = Val(txtHdcpKeySpec.Text)               'HDCPM
    rs.Fields(14) = Val(txtResolutionSpec.Text)            'ResolutionM
    rs.Fields(15) = Val(txtMacAddrSpec.Text)               'MACAddrM
    rs.Fields(16) = Val(txtPartitionVerSpec.Text)          'PartitionVerM
    rs.Fields(17) = Val(txtAreaSpec.Text)                  'AreaM
    rs.Fields(18) = Val(txtDeviceKeySpec.Text)             'DeviceKeyM

    If Check1.Value = 1 Then rs.Fields(19) = True
    ElseIf Check1.Value = 0 Then rs.Fields(19) = False
    End If
    If Check2.Value = 1 Then rs.Fields(20) = True
    ElseIf Check2.Value = 0 Then rs.Fields(20) = False
    End If
    If Check3.Value = 1 Then rs.Fields(21) = True
    ElseIf Check3.Value = 0 Then rs.Fields(21) = False
    End If
    If Check4.Value = 1 Then rs.Fields(22) = True
    ElseIf Check4.Value = 0 Then rs.Fields(22) = False
    End If
    If Check5.Value = 1 Then rs.Fields(23) = True
    ElseIf Check5.Value = 0 Then rs.Fields(23) = False
    End If
    If Check6.Value = 1 Then rs.Fields(24) = True
    ElseIf Check6.Value = 0 Then rs.Fields(24) = False
    End If
    If Check7.Value = 1 Then rs.Fields(26) = True
    ElseIf Check7.Value = 0 Then rs.Fields(26) = False
    End If
    If Check8.Value = 1 Then rs.Fields(27) = True
    ElseIf Check8.Value = 0 Then rs.Fields(27) = False
    End If
    If Check9.Value = 1 Then rs.Fields(28) = True
    ElseIf Check9.Value = 0 Then rs.Fields(28) = False
    End If
    If Check10.Value = 1 Then rs.Fields(30) = True
    ElseIf Check10.Value = 0 Then rs.Fields(30) = False
    End If
    If Check11.Value = 1 Then rs.Fields(31) = True
    ElseIf Check11.Value = 0 Then rs.Fields(31) = False
    End If
    If Check12.Value = 1 Then rs.Fields(32) = True
    ElseIf Check12.Value = 0 Then rs.Fields(32) = False
    End If
    If Check13.Value = 1 Then rs.Fields(25) = True
    ElseIf Check13.Value = 0 Then rs.Fields(25) = False
    End If
    If Check14.Value = 1 Then rs.Fields(29) = True
    ElseIf Check14.Value = 0 Then rs.Fields(29) = False
    End If
    If Check15.Value = 1 Then rs.Fields(33) = True
    ElseIf Check15.Value = 0 Then rs.Fields(33) = False
    End If
 
    rs.Update

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    MsgBox "Save success!", vbOKOnly, "warning"
    Unload Me
    Unload Form1

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

