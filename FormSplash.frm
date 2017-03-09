VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2205
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CommandLoadXml 
      Caption         =   "���������ļ�"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   2880
      Picture         =   "FormSplash.frx":1DF72
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   360
      Width           =   780
   End
   Begin VB.PictureBox PictureBrand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   360
      Picture         =   "FormSplash.frx":1FD66
      ScaleHeight     =   750
      ScaleWidth      =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2550
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   3240
      TabIndex        =   0
      Top             =   1920
      Width           =   570
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandLoadXml_Click()
    ' CancelError is True.
    On Error GoTo ErrHandler
    ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)|*.*|Xml Files (*.xml)|*.xml"
    ' Specify default filter.
    CommonDialog1.FilterIndex = 2

    ' Display the Open dialog box.
    CommonDialog1.ShowOpen
    gstrXmlPath = CommonDialog1.FileName
    
    Unload Me
    Exit Sub

ErrHandler:
    ' User pressed Cancel button.
    MsgBox "����� XML �ļ��������޷����������", vbExclamation, "���������ļ�"
    Exit Sub
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LoadFormMain
End Sub
