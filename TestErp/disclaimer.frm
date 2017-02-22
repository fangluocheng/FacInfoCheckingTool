VERSION 5.00
Begin VB.Form disclaimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disclaimer"
   ClientHeight    =   3390
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"disclaimer.frx":0000
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "disclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()

    disclaimer.Hide

End Sub
