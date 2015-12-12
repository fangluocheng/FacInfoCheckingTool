Attribute VB_Name = "Module3"
Option Explicit

Public Const AdjustSingle = 1
Public Const AdjustDouble = 0

Public Const SingleStep = 0
Public Const ComplexStep = 1

Public Const HighBri = 1
Public Const LowBri = 0

Type ColorTemp
    X As Single
    Y As Single
    lv As Single
End Type

Public strBuff As String

Public i As Integer
Public IsStepTime As Long

Public Const xxf = 1
Public Const xfyf = 2
Public Const yyf = 3
Public Const microStep = True
Public Const StepbyStep = False

Public IsBarcodeLen As Integer
Public IsFunctionAutoBri As Boolean
Public IsSensorLight As Boolean
Public IsSaveData As Boolean
Public IsCheckColorTemp  As Boolean

Public IsSendOffset As Boolean
Public IsAdjsutOffset As Boolean

Public strCurrentModelName As String
Public strDataVersion As String
Public IsStop As Boolean
Public IsACK As Boolean
Public SetTVCurrentComID As Integer
Public SetData As Integer
Public SetDay As Integer

Public IsSNWriteSuccess As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long

Public SetTVCurrentComBaud As Long

