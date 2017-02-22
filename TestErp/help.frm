VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   $"help.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "Submitting a Quotation to the Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   6495
   End
   Begin VB.Label Label6 
      Caption         =   $"help.frx":01DE
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Interrogating the Quotation Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   6495
   End
   Begin VB.Label Label4 
      Caption         =   $"help.frx":0358
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Server Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   $"help.frx":048A
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to the Soapuser.com Quotation Service Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()

    help.Hide

End Sub
