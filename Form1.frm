VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "π§≥ß–≈œ¢–£—Èπ§æﬂ"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   16905
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   16905
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   12000
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TxtReceive 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      Left            =   12600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   120
      Width           =   4185
   End
   Begin VB.Timer Timer1 
      Left            =   11520
      Top             =   240
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ãı¬Î"
      Height          =   1125
      Left            =   4320
      TabIndex        =   33
      Top             =   6720
      Width           =   8175
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   0
         Text            =   "123456789"
         Top             =   240
         Width           =   7905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "≤‚ ‘Ω·π˚"
      Height          =   1125
      Left            =   120
      TabIndex        =   32
      Top             =   6720
      Width           =   4095
      Begin VB.Label lbResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Checking"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TV –≈œ¢"
      Height          =   5700
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   12375
      Begin VB.TextBox txtDeviceKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "None"
         Top             =   5025
         Width           =   4000
      End
      Begin VB.TextBox txtMacAddr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtResolution 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtHdcpKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "None"
         Top             =   5025
         Width           =   4000
      End
      Begin VB.TextBox txtCarrier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPanelName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtTwoPointFourVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPartitionVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "None"
         Top             =   5025
         Width           =   4000
      End
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtDimension 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtHWVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtFlashInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtSysVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtModelInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Device Key"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   14
         Left            =   8190
         TabIndex        =   30
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAC µÿ÷∑"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   11
         Left            =   8190
         TabIndex        =   28
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HDCP"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4K/2K"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   10
         Left            =   4150
         TabIndex        =   24
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "«¯”Ú"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   13
         Left            =   4150
         TabIndex        =   22
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≤•øÿ∆ΩÃ®"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   8
         Left            =   8190
         TabIndex        =   20
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2.4G ∞Ê±æ"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∆¡–Õ∫≈"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   7
         Left            =   4150
         TabIndex        =   16
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∑÷«¯∞Ê±æ"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "«˛µ¿"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   5
         Left            =   8190
         TabIndex        =   12
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "”≤º˛∞Ê±æ"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   4
         Left            =   4150
         TabIndex        =   8
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flash –≈œ¢"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   8190
         TabIndex        =   6
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "œµÕ≥∞Ê±æ"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   4150
         TabIndex        =   4
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ª˙–Õ"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4000
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu vbSet 
      Caption         =   "…Ë÷√"
      Begin VB.Menu tbSetComPort 
         Caption         =   "…Ë÷√¥Æø⁄"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "…Ë÷√ ˝æ›πÊ∏Ò"
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
Dim IsAllDataMatch As Boolean

Private Sub Form_Load()
    i = 0
    SetTVCurrentComBaud = 115200

    StepTime = IsStepTime
    If StepTime < 500 Then
        StepTime = 500
    End If

    IsStop = False
    
    If isUartMode Then
        subInitComPort
    Else
        subInitNetwork
    End If
    
    subInitInterface

    Label1 = strCurrentModelName
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit
  
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
  
    End
Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub


Private Sub subInitInterface()
    txtInput.Text = ""
    txtInput.Locked = False
    isCmdDataRecv = False
    
    'Whether the CheckBox of database file(*.mdb) selected or not.
    'If not, config the TextBox
    If Not IsModelSelected Then
        txtModelInfo.Text = strChkBoxUnselected
        txtModelInfo.BackColor = &HE0E0E0
    End If
    If Not IsSysVerSelected Then
        txtSysVer.Text = strChkBoxUnselected
        txtSysVer.BackColor = &HE0E0E0
    End If
    If Not IsFlashInfoSelected Then
        txtFlashInfo.Text = strChkBoxUnselected
        txtFlashInfo.BackColor = &HE0E0E0
    End If
    If Not IsHardwareVerSelected Then
        txtHWVer.Text = strChkBoxUnselected
        txtHWVer.BackColor = &HE0E0E0
    End If
    If Not IsDimensionSelected Then
        txtDimension.Text = strChkBoxUnselected
        txtDimension.BackColor = &HE0E0E0
    End If
    If Not IsChannelSelected Then
        txtChannel.Text = strChkBoxUnselected
        txtChannel.BackColor = &HE0E0E0
    End If
    If Not Is24GVerSelected Then
        txtTwoPointFourVer.Text = strChkBoxUnselected
        txtTwoPointFourVer.BackColor = &HE0E0E0
    End If
    If Not IsPanelSelected Then
        txtPanelName.Text = strChkBoxUnselected
        txtPanelName.BackColor = &HE0E0E0
    End If
    If Not IsCarrierSelected Then
        txtCarrier.Text = strChkBoxUnselected
        txtCarrier.BackColor = &HE0E0E0
    End If
    If Not IsHDCPSelected Then
        txtHdcpKey.Text = strChkBoxUnselected
        txtHdcpKey.BackColor = &HE0E0E0
    End If
    If Not IsResolutionSelected Then
        txtResolution.Text = strChkBoxUnselected
        txtResolution.BackColor = &HE0E0E0
    End If
    If Not IsMACAddrSelected Then
        txtMacAddr.Text = strChkBoxUnselected
        txtMacAddr.BackColor = &HE0E0E0
    End If
    If Not IsPartitionVerSelected Then
        txtPartitionVer.Text = strChkBoxUnselected
        txtPartitionVer.BackColor = &HE0E0E0
    End If
    If Not IsAreaVerSelected Then
        txtArea.Text = strChkBoxUnselected
        txtArea.BackColor = &HE0E0E0
    End If
    If Not IsDeviceKeySelected Then
        txtDeviceKey.Text = strChkBoxUnselected
        txtDeviceKey.BackColor = &HE0E0E0
    End If
End Sub

Private Sub subInitComPort()
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    If rs.EOF = False Then
        SetTVCurrentComID = rs("ComID")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End
    End If

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    ComInit

End Sub

Private Sub subInitNetwork()
    isNetworkConnected = False
    With tcpClient
        .Protocol = sckTCPProtocol
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = strRemoteHost
        .RemotePort = lngRemotePort
    End With
End Sub

Private Function funSNWrite() As Boolean
    strSerialNo = ""
    scanbarcode = ""
    strSerialNo = UCase$(txtInput.Text)
    
    If subJudgeTheSNIsAvailable = True Then
        funSNWrite = True
        scanbarcode = strSerialNo
    Else
        funSNWrite = False
    End If
End Function

Private Sub subInitBeforeRunning()
    IsSNWriteSuccess = True
    IsAllDataMatch = True
    txtInput.Locked = True
    isCmdDataRecv = False
    strSerialNo = ""
    
    If IsModelSelected Then
        txtModelInfo.Text = strNoRecvData
        txtModelInfo.BackColor = &HFFFFFF
    End If
    If IsSysVerSelected Then
        txtSysVer.Text = strNoRecvData
        txtSysVer.BackColor = &HFFFFFF
    End If
    If IsFlashInfoSelected Then
        txtFlashInfo.Text = strNoRecvData
        txtFlashInfo.BackColor = &HFFFFFF
    End If
    If IsHardwareVerSelected Then
        txtHWVer.Text = strNoRecvData
        txtHWVer.BackColor = &HFFFFFF
    End If
    If IsDimensionSelected Then
        txtDimension.Text = strNoRecvData
        txtDimension.BackColor = &HFFFFFF
    End If
    If IsChannelSelected Then
        txtChannel.Text = strNoRecvData
        txtChannel.BackColor = &HFFFFFF
    End If
    If Is24GVerSelected Then
        txtTwoPointFourVer.Text = strNoRecvData
        txtTwoPointFourVer.BackColor = &HFFFFFF
    End If
    If IsPanelSelected Then
        txtPanelName.Text = strNoRecvData
        txtPanelName.BackColor = &HFFFFFF
    End If
    If IsCarrierSelected Then
        txtCarrier.Text = strNoRecvData
        txtCarrier.BackColor = &HFFFFFF
    End If
    If IsHDCPSelected Then
        txtHdcpKey.Text = strNoRecvData
        txtHdcpKey.BackColor = &HFFFFFF
    End If
    If IsResolutionSelected Then
        txtResolution.Text = strNoRecvData
        txtResolution.BackColor = &HFFFFFF
    End If
    If IsMACAddrSelected Then
        txtMacAddr.Text = strNoRecvData
        txtMacAddr.BackColor = &HFFFFFF
    End If
    If IsPartitionVerSelected Then
        txtPartitionVer.Text = strNoRecvData
        txtPartitionVer.BackColor = &HFFFFFF
    End If
    If IsAreaVerSelected Then
        txtArea.Text = strNoRecvData
        txtArea.BackColor = &HFFFFFF
    End If
    If IsDeviceKeySelected Then
        txtDeviceKey.Text = strNoRecvData
        txtDeviceKey.BackColor = &HFFFFFF
    End If
    
    lbResult.Caption = "Checking"
    lbResult.BackColor = &HFFFFFF
    Log_Clear
    TxtReceive.ForeColor = &H80000008
    lbResult.FontSize = 22
    
End Sub

Private Function subJudgeTheSNIsAvailable() As Boolean
    If strSerialNo = "" Or Len(strSerialNo) <> barcodeLen Then
        Log_Clear
        Log_Info "Please confirm the SN again?"
        txtInput.Text = ""
        txtInput.SetFocus
        subJudgeTheSNIsAvailable = False
    Else
        subJudgeTheSNIsAvailable = True
        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
End Function

Private Sub subInitAfterRunning()
    IsSNWriteSuccess = False
    txtInput.Locked = False
    txtInput.Text = ""
    txtInput.SetFocus
    
    isNetworkConnected = False
    tcpClient.Close
End Sub

Private Sub subMainProcesser()
    Dim i, j As Integer

On Error GoTo ErrExit
    subInitBeforeRunning
    If IsStop = True Then
        Exit Sub
    End If

    If IsSNWriteSuccess = funSNWrite Then
        If IsStop = True Then
            Exit Sub
        End If
        txtInput = scanbarcode
    Else
        'ShowError_Sys (6)
        GoTo FAIL
    End If

On Error GoTo ErrExit
    j = 0

RESEND_CMD_0:
    ClearComBuf
    'Send cmd, read data and save data
    'Enter factory mode fisrt, or other cmd may not respond.
    ENTER_FAC_MODE
    DelayMS StepTime
    Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
    If isCmdDataRecv = False Then
        If j > cmdResendTimes Then
            j = 0
            Log_Info "Cannot read enter factory. Please do the Letv Reset!!!"
            MsgBox "Please do the Letv Reset!"
            GoTo FAIL
        Else
            j = j + 1
            Log_Info "Resend cmd ENTER_FAC_MODE!!!"
            GoTo RESEND_CMD_0
        End If
    Else
        j = 0
        GoTo RESEND_CMD_1
    End If
    
RESEND_CMD_1:
    If IsModelSelected Then
        ClearComBuf
        READ_MODEL_NAME
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the modeal name!!!"
                GoTo RESEND_CMD_2
            Else
                j = j + 1
                Log_Info "Resend cmd READ_MODEL_NAME!!!"
                GoTo RESEND_CMD_1
            End If
        Else
            j = 0
            GoTo RESEND_CMD_2
        End If
    End If
    
RESEND_CMD_2:
    If IsSysVerSelected Then
        ClearComBuf
        READ_SYS_VERSION
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the system version!!!"
                GoTo RESEND_CMD_3
            Else
                j = j + 1
                Log_Info "Resend cmd READ_SYS_VERSION!!!"
                GoTo RESEND_CMD_2
            End If
        Else
            j = 0
            GoTo RESEND_CMD_3
        End If
    End If
    
RESEND_CMD_3:
    If IsFlashInfoSelected Then
        ClearComBuf
        READ_FLASH_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the Flash info!!!"
                GoTo RESEND_CMD_4
            Else
                j = j + 1
                Log_Info "Resend cmd READ_FLASH_INFO!!!"
                GoTo RESEND_CMD_3
            End If
        Else
            j = 0
            GoTo RESEND_CMD_4
        End If
    End If
    
RESEND_CMD_4:
    If IsHardwareVerSelected Then
        ClearComBuf
        READ_HARDWARE_VERSION
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the hardware version!!!"
                GoTo RESEND_CMD_5
            Else
                j = j + 1
                Log_Info "Resend cmd READ_HARDWARE_VERSION!!!"
                GoTo RESEND_CMD_4
            End If
        Else
            j = 0
            GoTo RESEND_CMD_5
        End If
    End If
    
RESEND_CMD_5:
    If IsDimensionSelected Then
        ClearComBuf
        READ_DIMENSION_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the dimension info!!!"
                GoTo RESEND_CMD_6
            Else
                j = j + 1
                Log_Info "Resend cmd READ_DIMENSION_INFO!!!"
                GoTo RESEND_CMD_5
            End If
        Else
            j = 0
            GoTo RESEND_CMD_6
        End If
    End If
    
RESEND_CMD_6:
    If IsChannelSelected Then
        ClearComBuf
        READ_CHANNEL_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the channel info!!!"
                GoTo RESEND_CMD_7
            Else
                j = j + 1
                Log_Info "Resend cmd READ_CHANNEL_INFO!!!"
                GoTo RESEND_CMD_6
            End If
        Else
            j = 0
            GoTo RESEND_CMD_7
        End If
    End If
    
RESEND_CMD_7:
    If Is24GVerSelected Then
        ClearComBuf
        READ_24G_VERSION
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the 2.4G version!!!"
                GoTo RESEND_CMD_8
            Else
                j = j + 1
                Log_Info "Resend cmd READ_24G_VERSION!!!"
                GoTo RESEND_CMD_7
            End If
        Else
            j = 0
            GoTo RESEND_CMD_8
        End If
    End If
    
RESEND_CMD_8:
    If IsPanelSelected Then
        ClearComBuf
        READ_PANEL_NAME
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the panel name!!!"
                GoTo RESEND_CMD_9
            Else
                j = j + 1
                Log_Info "Resend cmd READ_PANEL_NAME!!!"
                GoTo RESEND_CMD_8
            End If
        Else
            j = 0
            GoTo RESEND_CMD_9
        End If
    End If
    
RESEND_CMD_9:
    If IsCarrierSelected Then
        ClearComBuf
        READ_CARRIER_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the carrier info!!!"
                GoTo RESEND_CMD_10
            Else
                j = j + 1
                Log_Info "Resend cmd READ_CARRIER_INFO!!!"
                GoTo RESEND_CMD_9
            End If
        Else
            j = 0
            GoTo RESEND_CMD_10
        End If
    End If
    
RESEND_CMD_10:
    If IsHDCPSelected Then
        ClearComBuf
        READ_HDCP_KEY
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the HDCP Key!!!"
                GoTo RESEND_CMD_11
            Else
                j = j + 1
                Log_Info "Resend cmd READ_HDCP_KEY!!!"
                GoTo RESEND_CMD_10
            End If
        Else
            j = 0
            GoTo RESEND_CMD_11
        End If
    End If
    
RESEND_CMD_11:
    If IsResolutionSelected Then
        ClearComBuf
        READ_RESOLUTION_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the resolution info!!!"
                GoTo RESEND_CMD_12
            Else
                j = j + 1
                Log_Info "Resend cmd READ_RESOLUTION_INFO!!!"
                GoTo RESEND_CMD_11
            End If
        Else
            j = 0
            GoTo RESEND_CMD_12
        End If
    End If
    
RESEND_CMD_12:
    If IsMACAddrSelected Then
        ClearComBuf
        READ_MAC_ADDRESS
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the MAC Address!!!"
                GoTo RESEND_CMD_13
            Else
                j = j + 1
                Log_Info "Resend cmd READ_MAC_ADDRESS!!!"
                GoTo RESEND_CMD_12
            End If
        Else
            j = 0
            GoTo RESEND_CMD_13
        End If
    End If
    
RESEND_CMD_13:
    If IsPartitionVerSelected Then
        ClearComBuf
        READ_PARTITION_VER
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the partition version!!!"
                GoTo RESEND_CMD_14
            Else
                j = j + 1
                Log_Info "Resend cmd READ_PARTITION_VER!!!"
                GoTo RESEND_CMD_13
            End If
        Else
            j = 0
            GoTo RESEND_CMD_14
        End If
    End If
    
RESEND_CMD_14:
    If IsAreaVerSelected Then
        ClearComBuf
        READ_AREA_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the area info!!!"
                GoTo RESEND_CMD_15
            Else
                j = j + 1
                Log_Info "Resend cmd READ_AREA_INFO!!!"
                GoTo RESEND_CMD_14
            End If
        Else
            j = 0
            GoTo RESEND_CMD_15
        End If
    End If
    
RESEND_CMD_15:
    If IsDeviceKeySelected Then
        ClearComBuf
        READ_DEVICE_KEY
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the device key!!!"
                GoTo RESEND_CMD_16
            Else
                j = j + 1
                Log_Info "Resend cmd READ_DEVICE_KEY!!!"
                GoTo RESEND_CMD_15
            End If
        Else
            j = 0
            GoTo RESEND_CMD_16
        End If
    End If
    
RESEND_CMD_16:
    ClearComBuf
    'Either PASS or FAIL, send "Exit factory mode" cmd.
    EXIT_FAC_MODE
    DelayMS StepTime
   
    If txtModelInfo.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtModelInfo.BackColor = &HFF&
    End If
    
    If txtSysVer.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtSysVer.BackColor = &HFF&
    End If
    
    If txtFlashInfo.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtFlashInfo.BackColor = &HFF&
    End If
    
    If txtHWVer.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtHWVer.BackColor = &HFF&
    End If
    If txtDimension.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtDimension.BackColor = &HFF&
    End If
    
    If txtChannel.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtChannel.BackColor = &HFF&
    End If
    
    If txtPartitionVer.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtPartitionVer.BackColor = &HFF&
    End If
    
    If txtTwoPointFourVer.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtTwoPointFourVer.BackColor = &HFF&
    End If
    
    If txtPanelName.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtPanelName.BackColor = &HFF&
    End If
    
    If txtCarrier.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtCarrier.BackColor = &HFF&
    End If
    
    If txtArea.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtArea.BackColor = &HFF&
    End If
    
    If txtHdcpKey.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtHdcpKey.BackColor = &HFF&
    End If
    If txtResolution.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtResolution.BackColor = &HFF&
    End If
    
    If txtMacAddr.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtMacAddr.BackColor = &HFF&
    End If
    
    If txtDeviceKey.Text = strNoRecvData Then
        IsAllDataMatch = False
        txtDeviceKey.BackColor = &HFF&
    End If
    
    If Not IsAllDataMatch Then
        GoTo FAIL
    End If

    Call saveAllData
    
PASS:
    lbResult.Caption = "PASS"
    lbResult.BackColor = &HFF00&
    Call subInitAfterRunning
    
    Exit Sub

FAIL:
    lbResult.Caption = "NG"
    lbResult.BackColor = &HFF&
    Call subInitAfterRunning

    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source

End Sub


Private Sub saveAllData()

    If strSerialNo = "" Then
        Exit Sub
    Else
        sqlstring = "select * from DataRecord"
        Executesql (sqlstring)
        rs.AddNew

        rs.Fields(0) = strCurrentModelName
        rs.Fields(1) = strSerialNo

        rs.Fields(2) = txtModelInfo.Text
        rs.Fields(3) = txtSysVer.Text
        rs.Fields(4) = txtFlashInfo.Text
        rs.Fields(5) = txtHWVer.Text
        rs.Fields(6) = txtDimension.Text
        rs.Fields(7) = txtChannel.Text
        rs.Fields(8) = txtPartitionVer.Text
        rs.Fields(9) = txtTwoPointFourVer.Text
        rs.Fields(10) = txtPanelName.Text
        rs.Fields(11) = txtCarrier.Text
        rs.Fields(12) = txtArea.Text
        rs.Fields(13) = txtHdcpKey.Text
        rs.Fields(14) = txtResolution.Text
        rs.Fields(15) = txtMacAddr.Text
        rs.Fields(16) = txtDeviceKey.Text
        
        rs.Fields(17) = Date
        rs.Fields(18) = Time
        
        rs.Update
        
        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If

End Sub


Private Sub tbSetComPort_Click()
    Form2.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    i = 0
    'ASCII = 13 means "Enter" of keyboard.
    If KeyAscii = 13 Then
        IsStop = False
        isNetworkConnected = False
        
        If txtInput.Locked = False Then
            If isUartMode = True Then
                subMainProcesser
            Else
                Do
                    tcpClient.Connect
                    Call DelaySWithCmdFlag(cmdReceiveWaitS * 2, isNetworkConnected)
                
                    If isNetworkConnected = True Then
                        subMainProcesser
                        Exit Do
                    Else
                        tcpClient.Close
                        i = i + 1
                    End If
                    Log_Info "Re-connect to TV."
                Loop While i <= 5
            End If
        End If
         
        If IsStop = True Then
            Exit Sub
        End If
    End If
End Sub

Private Sub vbSetSPEC_Click()
    frmSetData.Show
End Sub


'------------------------------------------------------------------------------
' MSComm related function
'------------------------------------------------------------------------------
Private Sub MSComm1_OnComm()
    
On Error GoTo Err
    Select Case MSComm1.CommEvent
        Case comEvReceive
            DelayMS StepTime
            Call hexReceive
        'Case comEvSend
        Case Else
    End Select
Err:
  
End Sub

Private Sub hexReceive()
 
On Error GoTo Err
    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim Counter As Integer
    Dim i, tmp, firstByteOfDataIdx As Integer
    
    firstByteOfDataIdx = 0
    Counter = MSComm1.InBufferCount

    If (Counter > 0) Then
        receiveData = ""
        ReceiveArr = MSComm1.Input

        'Find ACK1 and ACK2, which are metioned in Letv's document.
        'ACK1 must be one of {0xC7, 0xCB, 0xCC, 0xD3, 0xD4, 0xD8, 0xDD, 0xE3, 0xE4, 0xE8, 0xED, 0xF0, 0xF5, 0xF9, 0xFE, 0xC2}
        For i = 0 To (Counter - 1) Step 1
            If i < (Counter - 1) Then
                If (ReceiveArr(i) Xor 255) = ReceiveArr(i + 1) Then
                    For Each tmp In Array(199, 203, 204, 211, 212, 216, 221, 227, 228, 232, 237, 240, 245, 249, 254, 194)
                        If ReceiveArr(i) = tmp Then
                            firstByteOfDataIdx = i
                            Exit For
                        End If
                    Next tmp
                    
                    If firstByteOfDataIdx = 14 Then
                        Exit For
                    Else
                        firstByteOfDataIdx = 0
                    End If
                End If
            End If
        Next i
        
        If firstByteOfDataIdx > 0 Then
            For i = firstByteOfDataIdx To (Counter - 1) Step 1
                If (ReceiveArr(i) < 16) Then
                    receiveData = receiveData & "0" & Hex(ReceiveArr(i)) & Space(1)
                Else
                    receiveData = receiveData & Hex(ReceiveArr(i)) & Space(1)
                End If
            Next i
        
            Log_Info receiveData
            
            'Update the CheckBoxes in the Form1.
            receiveData = ""
            
            'Data starts from ReceiveArr(firstByteOfDataIdx + 3). DataLength is ReceiveArr(firstByteOfDataIdx + 2).
            For i = (firstByteOfDataIdx + 3) To ((firstByteOfDataIdx + 3) + ReceiveArr(firstByteOfDataIdx + 2) - 1) Step 1
                If cmdIdentifyNum = 5 Or cmdIdentifyNum = 9 Or cmdIdentifyNum = 11 Or cmdIdentifyNum = 12 Then
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData & "0" & Hex(ReceiveArr(i))
                    Else
                        receiveData = receiveData & Hex(ReceiveArr(i))
                    End If
                Else
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData & "0" & Chr(ReceiveArr(i))
                    Else
                        receiveData = receiveData & Chr(ReceiveArr(i))
                    End If
                End If
            Next i
            
            Select Case cmdIdentifyNum
                Case 0
                    isCmdDataRecv = True
                Case 2                                     'System Version
                    isCmdDataRecv = True
                    If IsSysVerSelected Then
                        If SysVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtSysVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtSysVer.BackColor = &HFF&
                        End If
                        
                        txtSysVer.Text = receiveData
                    End If
                Case 3                                     'Flash Info
                    isCmdDataRecv = True
                    If IsFlashInfoSelected Then
                        txtFlashInfo.Text = receiveData & "G"
                        
                        If FlashInfoSpec = txtFlashInfo.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtFlashInfo.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtFlashInfo.BackColor = &HFF&
                        End If
                    End If
                Case 4                                     'Hardware Version
                    isCmdDataRecv = True
                    If IsHardwareVerSelected Then
                        If HardwareVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtHWVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtHWVer.BackColor = &HFF&
                        End If
                        
                        txtHWVer.Text = receiveData
                    End If
                Case 5                                     '3D\2D
                    isCmdDataRecv = True
                    If IsDimensionSelected Then
                        If receiveData = "00" Then
                            txtDimension.Text = "3D"
                        Else
                            txtDimension.Text = "2D"
                        End If
                        
                        If DimensionSpec = txtDimension.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtDimension.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtDimension.BackColor = &HFF&
                        End If
                    End If
                Case 6                                     '2.4G Version
                    isCmdDataRecv = True
                    If Is24GVerSelected Then
                        If TwoPointFourGVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtTwoPointFourVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtTwoPointFourVer.BackColor = &HFF&
                        End If
                        
                        txtTwoPointFourVer.Text = receiveData
                    End If
                Case 7                                     'Panel Name
                    isCmdDataRecv = True
                    If IsPanelSelected Then
                        If PanelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtPanelName.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtPanelName.BackColor = &HFF&
                        End If
                        
                        txtPanelName.Text = receiveData
                    End If
                Case 8                                     'Carrier Info
                    isCmdDataRecv = True
                    If IsCarrierSelected Then
                        If CarrierSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtCarrier.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtCarrier.BackColor = &HFF&
                        End If
                        
                        txtCarrier.Text = receiveData
                    End If
                Case 9                                     'HDCP Key
                    isCmdDataRecv = True
                    If IsHDCPSelected Then
                        'HDCP Key return 0x30 means HDCP Key is NOT written.
                        If receiveData = "30" Then
                            IsAllDataMatch = False
                            txtHdcpKey.BackColor = &HFF&
                            txtHdcpKey.Text = "HDCP Key Œ¥…’¬º"
                        Else
                            IsAllDataMatch = True And IsAllDataMatch
                            txtHdcpKey.BackColor = &HFF00&
                            txtHdcpKey.Text = "HDCP Key “—…’¬º"
                        End If
                    End If
                Case 10                                    'Model Name
                    isCmdDataRecv = True
                    If IsModelSelected Then
                        If ModelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtModelInfo.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtModelInfo.BackColor = &HFF&
                        End If
                        
                        txtModelInfo.Text = receiveData
                    End If
                Case 11                                    '4K\2K
                    isCmdDataRecv = True
                    If IsResolutionSelected Then
                        If receiveData = "00" Then
                            txtResolution.Text = "4K"
                        Else
                            txtResolution.Text = "2K"
                        End If
                        
                        If ResolutionSpec = txtResolution.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtResolution.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtResolution.BackColor = &HFF&
                        End If
                    End If
                Case 12                                    'MAC Address
                    isCmdDataRecv = True
                    If IsMACAddrSelected Then
                        If Len(receiveData) = 12 Then
                            sqlstring = "select * from DataRecord where MACAddr='" & receiveData & "'"
                            Executesql (sqlstring)
                            
                            If rs.RecordCount > 0 Then
                                If rs.RecordCount = 1 Then
                                    Log_Info "The MAC Address is the same as a TV whose SerialNO is [" & rs("SerialNO") & "]."
                                Else
                                    Log_Info "There are some TV whose MAC Address are the same. Please check the database file!!!"
                                End If
                                
                                TxtReceive.ForeColor = &HFF&
                                IsAllDataMatch = False
                                txtMacAddr.BackColor = &HFF&
                                txtMacAddr.Text = "MAC µÿ÷∑÷ÿ∏¥"
                            Else
                                IsAllDataMatch = True And IsAllDataMatch
                                txtMacAddr.BackColor = &HFF00&
                                txtMacAddr.Text = receiveData
                            End If
                            
                            Set cn = Nothing
                            Set rs = Nothing
                            sqlstring = ""
                        Else
                            Log_Info "The lenght of MAC address is wrong."
                            txtMacAddr.BackColor = &HFF&
                            txtMacAddr.Text = receiveData
                        End If
                    End If
                Case 13                                    'Channel Info
                    isCmdDataRecv = True
                    If IsChannelSelected Then
                        If ChannelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtChannel.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtChannel.BackColor = &HFF&
                        End If
                        
                        txtChannel.Text = receiveData
                    End If
                Case 14                                    'Partition Version
                    isCmdDataRecv = True
                    If IsPartitionVerSelected Then
                        If PartitionVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtPartitionVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtPartitionVer.BackColor = &HFF&
                        End If
                        
                        txtPartitionVer.Text = receiveData
                    End If
                Case 15                                    'Area Info
                    isCmdDataRecv = True
                    If IsAreaVerSelected Then
                        If AreaSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtArea.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtArea.BackColor = &HFF&
                        End If
                        
                        txtArea.Text = receiveData
                    End If
                Case 16                                    'Device Key
                    isCmdDataRecv = True
                    If IsDeviceKeySelected Then
                        If Len(receiveData) = 32 Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtDeviceKey.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtDeviceKey.BackColor = &HFF&
                        End If
                        
                        txtDeviceKey.Text = Strings.Right(receiveData, 5)
                    End If
                Case Else
                    Log_Info "Unknown command"
            End Select
        Else
            TxtReceive.Text = TxtReceive.Text & vbCrLf
            TxtReceive.SelStart = Len(TxtReceive.Text)
        End If
    Else
        'Ignore empty data
    End If
    
Err:

End Sub


Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Err
    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim i, tmp, firstByteOfDataIdx As Integer
    
    firstByteOfDataIdx = 0

    If (bytesTotal > 0) Then
        receiveData = ""
        tcpClient.GetData ReceiveArr, vbByte, bytesTotal
        
        If bytesTotal > 3 Then
            receiveData = ""
            
            For i = 3 To (bytesTotal - 1) Step 1
                If cmdIdentifyNum = 5 Or cmdIdentifyNum = 9 Or cmdIdentifyNum = 11 Or cmdIdentifyNum = 12 Then
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData & "0" & Hex(ReceiveArr(i))
                    Else
                        receiveData = receiveData & Hex(ReceiveArr(i))
                    End If
                Else
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData & "0" & Chr(ReceiveArr(i))
                    Else
                        receiveData = receiveData & Chr(ReceiveArr(i))
                    End If
                End If
            Next i
            
            Select Case cmdIdentifyNum
                Case 0
                    isCmdDataRecv = True
                Case 2                                     'System Version
                    isCmdDataRecv = True
                    If IsSysVerSelected Then
                        If SysVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtSysVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtSysVer.BackColor = &HFF&
                        End If
                        
                        txtSysVer.Text = receiveData
                    End If
                Case 3                                     'Flash Info
                    isCmdDataRecv = True
                    If IsFlashInfoSelected Then
                        txtFlashInfo.Text = receiveData & "G"
                        
                        If FlashInfoSpec = txtFlashInfo.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtFlashInfo.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtFlashInfo.BackColor = &HFF&
                        End If
                    End If
                Case 4                                     'Hardware Version
                    isCmdDataRecv = True
                    If IsHardwareVerSelected Then
                        If HardwareVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtHWVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtHWVer.BackColor = &HFF&
                        End If
                        
                        txtHWVer.Text = receiveData
                    End If
                Case 5                                     '3D\2D
                    isCmdDataRecv = True
                    If IsDimensionSelected Then
                        If receiveData = "00" Then
                            txtDimension.Text = "3D"
                        Else
                            txtDimension.Text = "2D"
                        End If
                        
                        If DimensionSpec = txtDimension.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtDimension.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtDimension.BackColor = &HFF&
                        End If
                    End If
                Case 6                                     '2.4G Version
                    isCmdDataRecv = True
                    If Is24GVerSelected Then
                        If TwoPointFourGVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtTwoPointFourVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtTwoPointFourVer.BackColor = &HFF&
                        End If
                        
                        txtTwoPointFourVer.Text = receiveData
                    End If
                Case 7                                     'Panel Name
                    isCmdDataRecv = True
                    If IsPanelSelected Then
                        If PanelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtPanelName.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtPanelName.BackColor = &HFF&
                        End If
                        
                        txtPanelName.Text = receiveData
                    End If
                Case 8                                     'Carrier Info
                    isCmdDataRecv = True
                    If IsCarrierSelected Then
                        If CarrierSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtCarrier.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtCarrier.BackColor = &HFF&
                        End If
                        
                        txtCarrier.Text = receiveData
                    End If
                Case 9                                     'HDCP Key
                    isCmdDataRecv = True
                    If IsHDCPSelected Then
                        'HDCP Key return 0x30 means HDCP Key is NOT written.
                        If receiveData = "30" Then
                            IsAllDataMatch = False
                            txtHdcpKey.BackColor = &HFF&
                            txtHdcpKey.Text = "HDCP Key Œ¥…’¬º"
                        Else
                            IsAllDataMatch = True And IsAllDataMatch
                            txtHdcpKey.BackColor = &HFF00&
                            txtHdcpKey.Text = "HDCP Key “—…’¬º"
                        End If
                    End If
                Case 10                                    'Model Name
                    isCmdDataRecv = True
                    If IsModelSelected Then
                        If ModelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtModelInfo.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtModelInfo.BackColor = &HFF&
                        End If
                        
                        txtModelInfo.Text = receiveData
                    End If
                Case 11                                    '4K\2K
                    isCmdDataRecv = True
                    If IsResolutionSelected Then
                        If receiveData = "00" Then
                            txtResolution.Text = "4K"
                        Else
                            txtResolution.Text = "2K"
                        End If
                        
                        If ResolutionSpec = txtResolution.Text Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtResolution.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtResolution.BackColor = &HFF&
                        End If
                    End If
                Case 12                                    'MAC Address
                    isCmdDataRecv = True
                    If IsMACAddrSelected Then
                        If Len(receiveData) = 12 Then
                            sqlstring = "select * from DataRecord where MACAddr='" & receiveData & "'"
                            Executesql (sqlstring)
                            
                            If rs.RecordCount > 0 Then
                                If rs.RecordCount = 1 Then
                                    Log_Info "The MAC Address is the same as a TV whose SerialNO is [" & rs("SerialNO") & "]."
                                Else
                                    Log_Info "There are some TV whose MAC Address are the same. Please check the database file!!!"
                                End If
                                
                                TxtReceive.ForeColor = &HFF&
                                IsAllDataMatch = False
                                txtMacAddr.BackColor = &HFF&
                                txtMacAddr.Text = "MAC µÿ÷∑÷ÿ∏¥"
                            Else
                                IsAllDataMatch = True And IsAllDataMatch
                                txtMacAddr.BackColor = &HFF00&
                                txtMacAddr.Text = receiveData
                            End If
                            
                            Set cn = Nothing
                            Set rs = Nothing
                            sqlstring = ""
                        Else
                            Log_Info "The lenght of MAC address is wrong."
                            txtMacAddr.BackColor = &HFF&
                            txtMacAddr.Text = receiveData
                        End If
                    End If
                Case 13                                    'Channel Info
                    isCmdDataRecv = True
                    If IsChannelSelected Then
                        If ChannelSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtChannel.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtChannel.BackColor = &HFF&
                        End If
                        
                        txtChannel.Text = receiveData
                    End If
                Case 14                                    'Partition Version
                    isCmdDataRecv = True
                    If IsPartitionVerSelected Then
                        If PartitionVerSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtPartitionVer.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtPartitionVer.BackColor = &HFF&
                        End If
                        
                        txtPartitionVer.Text = receiveData
                    End If
                Case 15                                    'Area Info
                    isCmdDataRecv = True
                    If IsAreaVerSelected Then
                        If AreaSpec = receiveData Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtArea.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtArea.BackColor = &HFF&
                        End If
                        
                        txtArea.Text = receiveData
                    End If
                Case 16                                    'Device Key
                    isCmdDataRecv = True
                    If IsDeviceKeySelected Then
                        If Len(receiveData) = 32 Then
                            IsAllDataMatch = True And IsAllDataMatch
                            txtDeviceKey.BackColor = &HFF00&
                        Else
                            IsAllDataMatch = False
                            txtDeviceKey.BackColor = &HFF&
                        End If
                        
                        txtDeviceKey.Text = Strings.Right(receiveData, 5)
                    End If
                Case Else
                    Log_Info "Unknown command"
            End Select
        Else
            TxtReceive.Text = TxtReceive.Text & vbCrLf
            TxtReceive.SelStart = Len(TxtReceive.Text)
        End If
    Else
        'Ignore empty data
    End If
    
Err:
End Sub

Private Sub tcpClient_Connect()
    'Success to connect the TV.
    isNetworkConnected = True
End Sub
