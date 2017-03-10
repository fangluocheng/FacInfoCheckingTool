VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "乐视属性比对工具"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14580
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "FormMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MAC地址条码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   42
      Top             =   6360
      Width           =   3735
      Begin VB.TextBox TextMacSN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Text            =   "123456789"
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox TxtReceive 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   6135
      Left            =   11040
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   120
      Width           =   3465
   End
   Begin VB.PictureBox PictureBrand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   120
      Picture         =   "FormMain.frx":1DF72
      ScaleHeight     =   750
      ScaleWidth      =   2520
      TabIndex        =   41
      Top             =   120
      Width           =   2550
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   10440
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   9960
      Top             =   240
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "整机条码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   6360
      Width           =   6975
      Begin VB.TextBox TextTvSN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Text            =   "123456789"
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "测试结果"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      TabIndex        =   17
      Top             =   6360
      Width           =   3495
      Begin VB.Label lbResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Checking"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10820
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   16
         Left            =   3650
         TabIndex        =   40
         Top             =   4760
         Width           =   3495
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Playready Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   16
         Left            =   3650
         TabIndex        =   39
         Top             =   4340
         Width           =   3495
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   4760
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Widevine Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   15
         Left            =   120
         TabIndex        =   37
         Top             =   4340
         Width           =   3495
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   14
         Left            =   7180
         TabIndex        =   36
         Top             =   3940
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   13
         Left            =   3650
         TabIndex        =   35
         Top             =   3940
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   3940
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   11
         Left            =   7180
         TabIndex        =   33
         Top             =   3120
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   10
         Left            =   3650
         TabIndex        =   32
         Top             =   3120
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   8
         Left            =   7180
         TabIndex        =   30
         Top             =   2300
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   7
         Left            =   3650
         TabIndex        =   29
         Top             =   2300
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   2300
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   5
         Left            =   7180
         TabIndex        =   27
         Top             =   1480
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   4
         Left            =   3650
         TabIndex        =   26
         Top             =   1480
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1480
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   2
         Left            =   7180
         TabIndex        =   24
         Top             =   660
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   1
         Left            =   3650
         TabIndex        =   23
         Top             =   660
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Device Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   14
         Left            =   7180
         TabIndex        =   16
         Top             =   3520
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "区域"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   11
         Left            =   7180
         TabIndex        =   15
         Top             =   2700
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "分区版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   2700
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4K/2K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   10
         Left            =   3650
         TabIndex        =   13
         Top             =   2700
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAC 地址"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   13
         Left            =   3650
         TabIndex        =   12
         Top             =   3520
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "播控平台"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   8
         Left            =   7180
         TabIndex        =   11
         Top             =   1880
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2.4G 版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1880
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "屏型号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   7
         Left            =   3650
         TabIndex        =   9
         Top             =   1880
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HDCP Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   12
         Left            =   120
         TabIndex        =   8
         Top             =   3520
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "渠道"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   5
         Left            =   7180
         TabIndex        =   7
         Top             =   1060
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "硬件版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1060
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   4
         Left            =   3650
         TabIndex        =   5
         Top             =   1060
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flash 信息"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   2
         Left            =   7180
         TabIndex        =   4
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "系统版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   1
         Left            =   3650
         TabIndex        =   3
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "机型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3500
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   2655
      TabIndex        =   19
      Top             =   120
      Width           =   8280
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsAllDataMatch As Boolean
Private mstrSNInput As String
Private mstrMacSn As String

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

Private Sub InitBeforeRunning()
    Dim i As Integer

    IsAllDataMatch = True
    gblCmdDataRecv = False
    mstrSNInput = ""
    
    For i = 0 To ITEMS_NUM
        If gutdPropertySetting.ItemChk(i) Then
            lbTVInfo(i).Caption = TVINFO_INIT
            lbTVInfo(i).BackColor = &HFFFFFF
        End If
    Next i
    
    lbResult.Caption = "Checking"
    lbResult.BackColor = &HFFFFFF
    Log_Clear
    TxtReceive.ForeColor = &H80000008
End Sub

Private Sub subInitAfterRunning()
    TextMacSN.Text = ""
    'TextMacSN.Enabled = False
    TextTvSN.Enabled = True
    TextTvSN.Text = ""
    TextTvSN.SetFocus
    
    If gblUartMode = False Then
        gblNetConnected = False
        tcpClient.Close
    End If
End Sub

Private Sub Run()
    On Error GoTo ErrExit
    Dim i, j As Integer

    InitBeforeRunning

    j = 0

RESEND_CMD_0:
    ClearComBuf
    'Send cmd, read data and save data
    'Enter factory mode fisrt, or other cmd may not respond.
    ENTER_FAC_MODE
    'DelayMS glngDelayTime
    Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)

    If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(0) Then
        ClearComBuf
        READ_MODEL_NAME
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)

        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(1) Then
        ClearComBuf
        READ_SYS_VERSION
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(2) Then
        ClearComBuf
        READ_FLASH_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(3) Then
        ClearComBuf
        READ_HARDWARE_VERSION
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(4) Then
        ClearComBuf
        READ_DIMENSION_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(5) Then
        ClearComBuf
        READ_CHANNEL_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(6) Then
        ClearComBuf
        READ_24G_VERSION
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(7) Then
        ClearComBuf
        READ_PANEL_NAME
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(8) Then
        ClearComBuf
        READ_CARRIER_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(9) Then
        ClearComBuf
        READ_PARTITION_VER
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the partition version!!!"
                GoTo RESEND_CMD_11
            Else
                j = j + 1
                Log_Info "Resend cmd READ_PARTITION_VER!!!"
                GoTo RESEND_CMD_10
            End If
        Else
            j = 0
            GoTo RESEND_CMD_11
        End If
    End If
    
RESEND_CMD_11:
    If gutdPropertySetting.ItemChk(10) Then
        ClearComBuf
        READ_RESOLUTION_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(11) Then
        ClearComBuf
        READ_AREA_INFO
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the area info!!!"
                GoTo RESEND_CMD_13
            Else
                j = j + 1
                Log_Info "Resend cmd READ_AREA_INFO!!!"
                GoTo RESEND_CMD_12
            End If
        Else
            j = 0
            GoTo RESEND_CMD_13
        End If
    End If
    
RESEND_CMD_13:
    If gutdPropertySetting.ItemChk(12) Then
        ClearComBuf
        READ_HDCP_KEY
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the HDCP Key!!!"
                GoTo RESEND_CMD_14
            Else
                j = j + 1
                Log_Info "Resend cmd READ_HDCP_KEY!!!"
                GoTo RESEND_CMD_13
            End If
        Else
            j = 0
            GoTo RESEND_CMD_14
        End If
    End If
    
RESEND_CMD_14:
    If gutdPropertySetting.ItemChk(13) Then
        ClearComBuf
        READ_MAC_ADDRESS
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the MAC Address!!!"
                GoTo RESEND_CMD_15
            Else
                j = j + 1
                Log_Info "Resend cmd READ_MAC_ADDRESS!!!"
                GoTo RESEND_CMD_14
            End If
        Else
            j = 0
            GoTo RESEND_CMD_15
        End If
    End If
    
RESEND_CMD_15:
    If gutdPropertySetting.ItemChk(14) Then
        ClearComBuf
        READ_DEVICE_KEY
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
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
    If gutdPropertySetting.ItemChk(15) Then
        ClearComBuf
        READ_WIDEVINE_KEY
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the Widevine key!!!"
                GoTo RESEND_CMD_17
            Else
                j = j + 1
                Log_Info "Resend cmd READ_WIDEVINE_KEY!!!"
                GoTo RESEND_CMD_16
            End If
        Else
            j = 0
            GoTo RESEND_CMD_17
        End If
    End If

RESEND_CMD_17:
    If gutdPropertySetting.ItemChk(16) Then
        ClearComBuf
        READ_PLAYREADY_KEY
        'DelayMS glngDelayTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, gblCmdDataRecv)
        
        If gblCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the Playready key!!!"
                GoTo RESEND_CMD_18
            Else
                j = j + 1
                Log_Info "Resend cmd READ_PLAYREADY_KEY!!!"
                GoTo RESEND_CMD_17
            End If
        Else
            j = 0
            GoTo RESEND_CMD_18
        End If
    End If

RESEND_CMD_18:
    ClearComBuf
    'Either PASS or FAIL, send "Exit factory mode" cmd.
    If gblExitFacCmd Then
        EXIT_FAC_MODE
        'DelayMS glngDelayTime
    End If

    For i = 0 To ITEMS_NUM
        If lbTVInfo(i).Caption = TVINFO_INIT Then
            IsAllDataMatch = False
            lbTVInfo(i).BackColor = &HFF&
        End If
    Next i
    
    If Not IsAllDataMatch Then
        GoTo FAIL
    End If

    'Call SaveData
    
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

Private Sub SaveData()
    Dim i As Integer

    If mstrSNInput = "" Then
        Exit Sub
    Else
        sqlstring = "select * from DataRecord"
        Executesql (sqlstring)
        rs.AddNew

        rs.Fields(0) = gutdPropertySetting.Items(0)
        rs.Fields(1) = mstrSNInput

        For i = 0 To ITEMS_NUM
            rs.Fields(i + 2) = lbTVInfo(i).Caption
        Next i
        
        rs.Fields(19) = Date
        rs.Fields(20) = Time
        
        rs.Update
        
        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If

End Sub

Private Sub TextTvSN_KeyPress(KeyAscii As Integer)
    Dim strTvSn As String
    
    If KeyAscii = 13 Then
        strTvSn = Trim$(TextTvSN.Text)
        If strTvSn = "" Or Len(strTvSn) <> gintSNLen Then
            Log_Clear
            Log_Info "Please confirm the SN again?"
            TextTvSN.Enabled = True
            TextTvSN.Text = ""
            TextTvSN.SetFocus
            MsgBox "输入的整机条码长度不对，请确认 XML 文件中设置的是否正确。", vbExclamation
            GoTo FAIL
        Else
            TextTvSN.Enabled = False
            TextMacSN.Enabled = True
            TextMacSN.SetFocus
        End If
    End If
    Exit Sub

FAIL:
    lbResult.Caption = "NG"
    lbResult.BackColor = &HFF&
End Sub

Private Sub TextMacSN_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrExit
    Dim i As Integer
    
    i = 0
    'ASCII = 13 means "Enter" of keyboard.
    If KeyAscii = 13 Then
        mstrMacSn = Trim$(TextMacSN.Text)
        If mstrMacSn = "" Or Len(mstrMacSn) <> gintMACLen Then
            Log_Clear
            Log_Info "Please confirm the MAC again?"
            TextMacSN.Enabled = True
            TextMacSN.Text = ""
            MsgBox "输入的 MAC 地址长度不对，请确认 XML 文件中设置的是否正确。", vbExclamation
            GoTo FAIL
        Else
            TextMacSN.Enabled = False

            If gblUartMode = True Then
                If MSComm1.PortOpen = False Then
                    MSComm1.PortOpen = True
                End If
                Run
            Else
                gblNetConnected = False
                Do
                    If tcpClient.State = sckClosed Then
                        Log_Info "TCP Connect"
                        tcpClient.Connect
                    End If
                    Call DelaySWithCmdFlag(cmdReceiveWaitS * 2, gblNetConnected)

                    If tcpClient.State = sckConnected Then
                        Run
                        Exit Do
                    Else
                        If tcpClient.State <> sckClosed Then
                            tcpClient.Close
                        End If
                        i = i + 1
                    End If
                    Log_Info "Re-connect to TV."
                Loop While i <= 5
                TextMacSN.Enabled = True
                TextMacSN.Text = ""
            End If
        End If
    End If
    Exit Sub

FAIL:
    lbResult.Caption = "NG"
    lbResult.BackColor = &HFF&
    Exit Sub

ErrExit:
    'Invalid Port Number
    If Err.Number = 8002 Then
        TextMacSN.Text = ""
    End If
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub MSComm1_OnComm()
    On Error GoTo Err

    Select Case MSComm1.CommEvent
        Case comEvReceive
            DelayMS glngDelayTime
            Call hexReceive
        'Case comEvSend
        Case Else
    End Select
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub hexReceive()
    On Error GoTo Err

    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim Counter As Integer
    Dim i, tmp, firstByteOfDataIdx As Integer
    
    firstByteOfDataIdx = 0
    Counter = MSComm1.InBufferCount

    'Ignore empty data
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
            
            'Update the CheckBoxes in the FormMain.
            receiveData = ""
            
            'Data starts from ReceiveArr(firstByteOfDataIdx + 3). DataLength is ReceiveArr(firstByteOfDataIdx + 2).
            For i = (firstByteOfDataIdx + 3) To ((firstByteOfDataIdx + 3) + ReceiveArr(firstByteOfDataIdx + 2) - 1) Step 1
                If gintCmdId = 6 Or gintCmdId = 12 Or _
                    gintCmdId = 14 Or gintCmdId = 15 Or _
                    gintCmdId = 17 Or gintCmdId = 18 Then
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
            
            InfoCompare gintCmdId, receiveData
        Else
            TxtReceive.Text = TxtReceive.Text & vbCrLf
            TxtReceive.SelStart = Len(TxtReceive.Text)
        End If
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub


Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo Err

    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim i, tmp, firstByteOfDataIdx As Integer
    
    firstByteOfDataIdx = 0

    'Ignore empty data
    If (bytesTotal > 0) Then
        receiveData = ""
        tcpClient.GetData ReceiveArr, vbByte, bytesTotal
        
        For i = 0 To (bytesTotal - 1) Step 1
            If (ReceiveArr(i) < 16) Then
                receiveData = receiveData & "0" & Hex(ReceiveArr(i)) & Space(1)
            Else
                receiveData = receiveData & Hex(ReceiveArr(i)) & Space(1)
            End If
        Next i
        
        Log_Info receiveData
        
        If bytesTotal >= 3 Then
            receiveData = ""
            
            For i = 3 To (bytesTotal - 1) Step 1
                If gintCmdId = 6 Or gintCmdId = 12 Or _
                    gintCmdId = 14 Or gintCmdId = 15 Or _
                    gintCmdId = 17 Or gintCmdId = 18 Then
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
            
            InfoCompare gintCmdId, receiveData
        Else
            TxtReceive.Text = TxtReceive.Text & vbCrLf
            TxtReceive.SelStart = Len(TxtReceive.Text)
        End If
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub tcpClient_Connect()
    'Success to connect the TV.
    gblNetConnected = True
End Sub

Private Sub InfoCompare(cmdIdx As Integer, recvData As String)
    Dim i As Integer

    For i = 0 To 16
        gblCmdDataRecv = True

        If cmdIdx = (i + 2) Then
            If cmdIdx = 4 Then                             '2.4G Version
                If gutdPropertySetting.ItemChk(2) Then
                    lbTVInfo(2).Caption = recvData & "G"
                            
                    If gutdPropertySetting.Items(2) = lbTVInfo(2).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(2).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(2).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 6 Then                         '2D or 3D
                If gutdPropertySetting.ItemChk(4) Then
                    If recvData = "00" Then
                        lbTVInfo(4).Caption = "3D"
                    Else
                        lbTVInfo(4).Caption = "2D"
                    End If
                            
                    If gutdPropertySetting.Items(4) = lbTVInfo(4).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(4).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(4).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 12 Then                        '4K or 2K
                If gutdPropertySetting.ItemChk(10) Then
                    If recvData = "00" Then
                        lbTVInfo(10).Caption = "4K"
                    Else
                        lbTVInfo(10).Caption = "2K"
                    End If
                            
                    If gutdPropertySetting.Items(10) = lbTVInfo(10).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(10).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(10).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 14 Then                        'HDCP Key
                If gutdPropertySetting.ItemChk(12) Then
                    'HDCP Key return 0x30 means HDCP Key is NOT written.
                    If recvData = "30" Then
                        IsAllDataMatch = False
                        lbTVInfo(12).BackColor = &HFF&
                        lbTVInfo(12).Caption = "HDCP Key 未烧录"
                    Else
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(12).BackColor = &HFF00&
                        lbTVInfo(12).Caption = "HDCP Key 已烧录"
                    End If
                End If
            ElseIf cmdIdx = 15 Then                        'MAC Address
                If gutdPropertySetting.ItemChk(13) Then
                    If Len(recvData) = gintMACLen Then
                        If UCase(recvData) = UCase(mstrMacSn) Then
                            IsAllDataMatch = True And IsAllDataMatch
                            lbTVInfo(13).BackColor = &HFF00&
                        Else
                            TxtReceive.ForeColor = &HFF&
                            IsAllDataMatch = False
                            lbTVInfo(13).BackColor = &HFF&
                        End If
                        lbTVInfo(13).Caption = recvData
                    Else
                        Log_Info "The lenght of MAC address is wrong."
                        lbTVInfo(13).BackColor = &HFF&
                        lbTVInfo(13).Caption = recvData
                    End If
                End If
            ElseIf cmdIdx = 16 Then                        'Device Key
                gblCmdDataRecv = True
                If gutdPropertySetting.ItemChk(14) Then
                    lbTVInfo(14).Caption = Strings.Right(recvData, 5)

                    If lbTVInfo(14).Caption = Strings.Right(Trim(TextTvSN.Text), 5) _
                    And Len(Trim(TextTvSN.Text)) >= 5 Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(14).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(14).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 17 Then
                gblCmdDataRecv = True
                If gutdPropertySetting.ItemChk(15) Then
                    'Widevine Key return 0x01 means Widevine Key is written.
                    If recvData = "01" Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(15).BackColor = &HFF00&
                        lbTVInfo(15).Caption = "Widevine Key 已烧录"
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(15).BackColor = &HFF&
                        lbTVInfo(15).Caption = "Widevine Key 未烧录"
                    End If
                End If
            ElseIf cmdIdx = 18 Then
                gblCmdDataRecv = True
                If gutdPropertySetting.ItemChk(16) Then
                    'Playready Key return 0x01 means Playready Key is written.
                    If recvData = "01" Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(16).BackColor = &HFF00&
                        lbTVInfo(16).Caption = "Playready Key 已烧录"
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(16).BackColor = &HFF&
                        lbTVInfo(16).Caption = "Playready Key 未烧录"
                    End If
                End If
            Else
                If gutdPropertySetting.ItemChk(i) Then
                    If gutdPropertySetting.Items(i) = recvData Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(i).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(i).BackColor = &HFF&
                    End If
                            
                    lbTVInfo(i).Caption = recvData
                End If
            End If
        End If
    Next i
End Sub

Private Sub ClearComBuf()
    If gblUartMode Then
        MSComm1.InBufferCount = 0
        MSComm1.OutBufferCount = 0
    End If
End Sub
