VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "²ÎÊýÉèÖÃ"
   ClientHeight    =   6990
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
   ScaleHeight     =   6990
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
      Top             =   120
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   5880
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
         Left            =   4560
         TabIndex        =   1
         Text            =   "500"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text4 
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
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()

 sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
Executesql (sqlstring)

 Label1 = strCurrentModelName
 
Text1.Text = rs("ComBaud")
Text2.Text = rs("Channel")
Text3.Text = rs("Delayms")
Text4.Text = rs("SN_Len")
Text5.Text = rs("WhitePattern")

If rs("COOL_2") Then
  Check1.Value = 1
Else
  Check1.Value = 0
End If
If rs("COOL_1") Then
  Check2.Value = 1
Else
  Check2.Value = 0
End If
If rs("NORMAL") Then
  Check3.Value = 1
Else
  Check3.Value = 0
End If
If rs("WARM_1") Then
  Check4.Value = 1
Else
  Check4.Value = 0
End If
If rs("WARM_2") Then
  Check5.Value = 1
Else
  Check5.Value = 0
End If


If rs("SaveData") Then
  Check6.Value = 1
Else
  Check6.Value = 0
End If
If rs("CheckColor") Then
  Check7.Value = 1
Else
  Check7.Value = 0
End If
If rs("SendOFF") Then
  Check8.Value = 1
Else
  Check8.Value = 0
End If
If rs("AdjustOFF") Then
  Check9.Value = 1
Else
  Check9.Value = 0
End If
If rs("SensorL") Then
  Check10.Value = 1
Else
  Check10.Value = 0
End If


Text7.Text = rs("Cool_1MI")
Text6.Text = rs("Cool_2MI")
Text8.Text = rs("NormalMI")
Text9.Text = rs("Warm_1MI")
Text10.Text = rs("Warm_2MI")


Set rs = Nothing
Set cn = Nothing
sqlstring = ""

End Sub

Private Sub Command1_Click()
  sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
Executesql (sqlstring)
  rs.Fields(1) = Val(Text1.Text)
  rs.Fields(2) = Val(Text2.Text)
  rs.Fields(3) = Val(Text3.Text)
  rs.Fields(4) = Val(Text5.Text)
  

  If Check1.Value = 1 Then rs.Fields(5) = True
  If Check1.Value = 0 Then rs.Fields(5) = False
  If Check2.Value = 1 Then rs.Fields(6) = True
  If Check2.Value = 0 Then rs.Fields(6) = False
  If Check3.Value = 1 Then rs.Fields(7) = True
  If Check3.Value = 0 Then rs.Fields(7) = False
  If Check4.Value = 1 Then rs.Fields(8) = True
  If Check4.Value = 0 Then rs.Fields(8) = False
  If Check5.Value = 1 Then rs.Fields(9) = True
  If Check5.Value = 0 Then rs.Fields(9) = False
  
  rs.Fields(10) = Val(Text4.Text)
 
  If Check6.Value = 1 Then rs.Fields(11) = True
  If Check6.Value = 0 Then rs.Fields(11) = False
  If Check7.Value = 1 Then rs.Fields(12) = True
  If Check7.Value = 0 Then rs.Fields(12) = False
  If Check8.Value = 1 Then rs.Fields(13) = True
  If Check8.Value = 0 Then rs.Fields(13) = False
  If Check9.Value = 1 Then rs.Fields(14) = True
  If Check9.Value = 0 Then rs.Fields(14) = False
  If Check10.Value = 1 Then rs.Fields(15) = True
  If Check10.Value = 0 Then rs.Fields(15) = False
  

  rs.Fields(16) = Val(Text6.Text)
  rs.Fields(17) = Val(Text7.Text)
  rs.Fields(18) = Val(Text8.Text)
  rs.Fields(19) = Val(Text9.Text)
  rs.Fields(20) = Val(Text10.Text)
  
  rs.Update

 Set cn = Nothing
 Set rs = Nothing
 sqlstring = ""

MsgBox "Save success!", vbOKOnly, "warning"
Unload Me
Unload Form1

End Sub

Private Sub Label4_Click()

End Sub
