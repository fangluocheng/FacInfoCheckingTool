VERSION 5.00
Begin VB.Form frmSetData 
   Caption         =   "设置条码长度"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form3"
   ScaleHeight     =   1575
   ScaleWidth      =   3390
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "条形码长度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    'Text1.Text = rs("SN_Len")
    
    'Set rs = Nothing
    'Set cn = Nothing

End Sub


Private Sub Command1_Click()
    'rs.Fields(10) = Val(Text1.Text)
    'rs.Update

    'Set cn = Nothing
    'Set rs = Nothing
    
    'rs.Update

    Unload Me
    Unload Form1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
