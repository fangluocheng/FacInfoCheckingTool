VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "工厂信息校验工具"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14730
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   10440
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TxtReceive 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   11040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   120
      Width           =   3585
   End
   Begin VB.Timer Timer1 
      Left            =   9960
      Top             =   240
   End
   Begin VB.Frame Frame3 
      Caption         =   "条码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   3840
      TabIndex        =   18
      Top             =   6360
      Width           =   7095
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   0
         Text            =   "123456789"
         Top             =   240
         Width           =   6850
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   850
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   3615
      Begin VB.Label lbResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Checking"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   20
         Top             =   240
         Width           =   3350
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu vbSet 
      Caption         =   "设置"
      Begin VB.Menu tbSetComPort 
         Caption         =   "设置串口"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "设置数据规格"
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

    StepTime = IsStepTime
    If StepTime < 500 Then
        StepTime = 500
    End If

    IsStop = False
    
    If isUartMode Then
        tbSetComPort.Enabled = True
        subInitComPort
    Else
        tbSetComPort.Enabled = False
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
    For i = 0 To itemNumOfTvInfo
        If Not chkTitleFlag(i) Then
            lbTVInfo(i).Caption = strChkBoxUnselected
            lbTVInfo(i).BackColor = &HE0E0E0
        End If
    Next i
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
    
    For i = 0 To itemNumOfTvInfo
        If chkTitleFlag(i) Then
            lbTVInfo(i).Caption = strNoRecvData
            lbTVInfo(i).BackColor = &HFFFFFF
        End If
    Next i
    
    lbResult.Caption = "Checking"
    lbResult.BackColor = &HFFFFFF
    Log_Clear
    TxtReceive.ForeColor = &H80000008
    'lbResult.FontSize = 22
    
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
    
    If isUartMode = False Then
        isNetworkConnected = False
        tcpClient.Close
    End If
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
        txtInput.Text = scanbarcode
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
    If chkTitleFlag(0) Then
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
    If chkTitleFlag(1) Then
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
    If chkTitleFlag(2) Then
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
    If chkTitleFlag(3) Then
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
    If chkTitleFlag(4) Then
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
    If chkTitleFlag(5) Then
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
    If chkTitleFlag(6) Then
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
    If chkTitleFlag(7) Then
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
    If chkTitleFlag(8) Then
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
    If chkTitleFlag(9) Then
        ClearComBuf
        READ_PARTITION_VER
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    If chkTitleFlag(10) Then
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
    If chkTitleFlag(11) Then
        ClearComBuf
        READ_AREA_INFO
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    If chkTitleFlag(12) Then
        ClearComBuf
        READ_HDCP_KEY
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    If chkTitleFlag(13) Then
        ClearComBuf
        READ_MAC_ADDRESS
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    If chkTitleFlag(14) Then
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
    If chkTitleFlag(15) Then
        ClearComBuf
        READ_WIDEVINE_KEY
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    If chkTitleFlag(16) Then
        ClearComBuf
        READ_PLAYREADY_KEY
        DelayMS StepTime
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
        
        If isCmdDataRecv = False Then
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
    EXIT_FAC_MODE
    DelayMS StepTime

    For i = 0 To itemNumOfTvInfo
        If lbTVInfo(i).Caption = strNoRecvData Then
            IsAllDataMatch = False
            lbTVInfo(i).BackColor = &HFF&
        End If
    Next i
    
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

        For i = 0 To itemNumOfTvInfo
            rs.Fields(i + 2) = lbTVInfo(i).Caption
        Next i
        
        rs.Fields(17) = Date
        rs.Fields(18) = Time
        
        rs.Update
        
        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If

End Sub


Private Sub tbSetComPort_Click()
    frmComPort.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    i = 0
    'ASCII = 13 means "Enter" of keyboard.
    If KeyAscii = 13 Then
        IsStop = False
        
        If txtInput.Locked = False Then
            If isUartMode = True Then
                subMainProcesser
            Else
                isNetworkConnected = False
                Do
                    If tcpClient.State = sckClosed Then
                        Log_Info "TCP Connect"
                        tcpClient.Connect
                        txtInput.Locked = True
                    End If
                    Call DelaySWithCmdFlag(cmdReceiveWaitS * 2, isNetworkConnected)

                    If tcpClient.State = sckConnected Then
                        subMainProcesser
                        Exit Do
                    Else
                        If tcpClient.State <> sckClosed Then
                            tcpClient.Close
                        End If
                        i = i + 1
                    End If
                    Log_Info "Re-connect to TV."
                Loop While i <= 5
                txtInput.Locked = False
                txtInput.Text = ""
                'MsgBox "Please connect TV and PC by network." & vbCrLf & _
                '    "Set PC IP to 192.168.1.2"
            End If
        End If
         
        If IsStop = True Then
            Exit Sub
        End If
    End If
    Exit Sub
    
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
                If cmdIdentifyNum = 6 Or cmdIdentifyNum = 12 Or _
                    cmdIdentifyNum = 14 Or cmdIdentifyNum = 15 Or _
                    cmdIdentifyNum = 17 Or cmdIdentifyNum = 18 Then
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
            
            InfoCompare cmdIdentifyNum, receiveData
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
                If cmdIdentifyNum = 6 Or cmdIdentifyNum = 12 Or _
                    cmdIdentifyNum = 14 Or cmdIdentifyNum = 15 Or _
                    cmdIdentifyNum = 17 Or cmdIdentifyNum = 18 Then
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
            
            InfoCompare cmdIdentifyNum, receiveData
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

Private Sub InfoCompare(cmdIdx As Integer, recvData As String)
    For i = 0 To 16
        isCmdDataRecv = True

        If cmdIdx = (i + 2) Then
            If cmdIdx = 4 Then                             '2.4G Version
                If chkTitleFlag(2) Then
                    lbTVInfo(2).Caption = recvData & "G"
                            
                    If strTvInfoSpec(2) = lbTVInfo(2).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(2).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(2).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 6 Then                         '2D or 3D
                If chkTitleFlag(4) Then
                    If recvData = "00" Then
                        lbTVInfo(4).Caption = "3D"
                    Else
                        lbTVInfo(4).Caption = "2D"
                    End If
                            
                    If strTvInfoSpec(4) = lbTVInfo(4).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(4).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(4).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 12 Then                        '4K or 2K
                If chkTitleFlag(10) Then
                    If recvData = "00" Then
                        lbTVInfo(10).Caption = "4K"
                    Else
                        lbTVInfo(10).Caption = "2K"
                    End If
                            
                    If strTvInfoSpec(10) = lbTVInfo(10).Caption Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(10).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(10).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 14 Then                        'HDCP Key
                If chkTitleFlag(12) Then
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
                If chkTitleFlag(13) Then
                    If Len(recvData) = 12 Then
                        sqlstring = "select * from DataRecord where MACAddr='" & recvData & "'"
                        Executesql (sqlstring)
                                
                        If rs.RecordCount > 0 Then
                            If rs.RecordCount = 1 Then
                                Log_Clear
                                Log_Info "请检查此电视机的条码是否为 [" & rs("SerialNO") & "]."
                            Else
                                Log_Info "在数据库中发现有多个相同的 MAC 地址，请检查"
                            End If
                                    
                            TxtReceive.ForeColor = &HFF&
                            IsAllDataMatch = False
                            lbTVInfo(13).BackColor = &HFF&
                            lbTVInfo(13).Caption = "MAC 地址重复"
                        Else
                            IsAllDataMatch = True And IsAllDataMatch
                            lbTVInfo(13).BackColor = &HFF00&
                            lbTVInfo(13).Caption = recvData
                        End If
                                
                        Set cn = Nothing
                        Set rs = Nothing
                        sqlstring = ""
                    Else
                        Log_Info "The lenght of MAC address is wrong."
                        lbTVInfo(13).BackColor = &HFF&
                        lbTVInfo(13).Caption = recvData
                    End If
                End If
            ElseIf cmdIdx = 16 Then                        'Device Key
                isCmdDataRecv = True
                If chkTitleFlag(14) Then
                    lbTVInfo(14).Caption = Strings.Right(recvData, 5)

                    If lbTVInfo(14).Caption = Strings.Right(Trim(txtInput.Text), 5) _
                    And Len(Trim(txtInput.Text)) >= 5 Then
                        IsAllDataMatch = True And IsAllDataMatch
                        lbTVInfo(14).BackColor = &HFF00&
                    Else
                        IsAllDataMatch = False
                        lbTVInfo(14).BackColor = &HFF&
                    End If
                End If
            ElseIf cmdIdx = 17 Then
                isCmdDataRecv = True
                If chkTitleFlag(15) Then
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
                isCmdDataRecv = True
                If chkTitleFlag(16) Then
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
                If chkTitleFlag(i) Then
                    If strTvInfoSpec(i) = recvData Then
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
