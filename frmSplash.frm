VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2580
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOpenExcelFile 
      BackColor       =   &H00808080&
      Caption         =   "机卡绑定机型"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog comDialogOpenExcelFile 
      Left            =   360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbModelName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "乐视工厂信息校验工具"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3660
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "请选择机型:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ss As Boolean

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrExit
    ss = False

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    sqlstring = "select * from CheckItem"
    Executesql (sqlstring)

    If rs.EOF = False Then
        rs.MoveFirst
        cmbModelName.Clear

        Do While Not rs.EOF
            cmbModelName.AddItem rs.Fields("Mark")
            rs.MoveNext
        Loop
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
        End
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)
    
    If rs.EOF = False Then
        strCurrentModelName = rs("CurrentModelName")
        SetTVCurrentComID = rs("ComID")
        SetData = rs("Date")
        SetDay = rs("Day")
        If rs("CommunicationMode") = "UART" Then
            isUartMode = True
            frmSetData.optUart.Value = True
            frmSetData.optNetwork.Value = False
        Else
            isUartMode = False
            frmSetData.optUart.Value = False
            frmSetData.optNetwork.Value = True
        End If
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End If
    
    Set cn = Nothing
    Set rs = Nothing

    sqlstring = ""
    cmbModelName = strCurrentModelName

    If SetData <> Day(Date) Then
        sqlstring = "select * from CommonTable where Mark='ATS'"
        Executesql (sqlstring)
        rs.Fields(4) = Day(Date)
        rs.Fields(5) = SetDay + 1
        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
    Exit Sub
    
ErrExit:
       MsgBox Err.Description, vbCritical, Err.Source
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit

    strCurrentModelName = cmbModelName
    sqlstring = ""
    sqlstring = "update CommonTable set CurrentModelName='" & strCurrentModelName & "' where Mark='ATS'"
    Executesql (sqlstring)
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from CheckItem where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    SetTVCurrentComBaud = rs("ComBaud")
    IsStepTime = rs("Delayms")
    barcodeLen = rs("SN_Len")
    
    For i = 0 To itemNumOfTvInfo
        chkTitleFlag(i) = rs.Fields(i + 16)
    Next i
    
    For i = 0 To 11
        strTvInfoSpec(i) = rs.Fields(i + 4)
    Next i

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

    If chkOpenExcelFile.Value = 0 Then
        isOpenDataBindingExcelFile = False
        Form1.Show
    Else
        isOpenDataBindingExcelFile = True
        ' Set filters.
        comDialogOpenExcelFile.Filter = "Excel Files 97-2003 (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx"
        ' Specify default filter.
        comDialogOpenExcelFile.FilterIndex = 2

        comDialogOpenExcelFile.ShowOpen
        
        strDataBindingExcelFileName = comDialogOpenExcelFile.FileName
        Form1.Show
    End If

    Exit Sub
    
ErrExit:
    MsgBox ("The Licence Key is Wrong.")
    
End Sub
