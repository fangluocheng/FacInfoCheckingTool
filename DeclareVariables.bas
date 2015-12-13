Attribute VB_Name = "Module3"
Option Explicit

Public Const AdjustSingle = 1
Public Const AdjustDouble = 0

Public Const SingleStep = 0
Public Const ComplexStep = 1

Public Const HighBri = 1
Public Const LowBri = 0

Public strBuff As String

Public i As Integer

Public Const xxf = 1
Public Const xfyf = 2
Public Const yyf = 3
Public Const microStep = True
Public Const StepbyStep = False

'------------------------------------------------------------------------------
'    Variable mapping the check items in "CheckItem".
'------------------------------------------------------------------------------
Public IsModel As Boolean
Public IsSysVer As Boolean
Public IsFlashInfo As Boolean
Public IsHardwareVer As Boolean
Public IsDimension As Boolean
Public IsChannel As Boolean
Public IsPartitionVer As Boolean
Public Is24GVer As Boolean
Public IsPanel As Boolean
Public IsCarrier As Boolean
Public IsArea As Boolean
Public IsHDCP As Boolean
Public IsResolution As Boolean
Public IsMACAddr As Boolean
Public IsDeviceKey As Boolean

'------------------------------------------------------------------------------
'    Variable mapping the items in "CheckItem".
'------------------------------------------------------------------------------
Public SetTVCurrentComBaud As Long                         'ComBaud
Public IsStepTime As Long                                  'Delayms
Public IsBarcodeLen As Integer                             'SN_Len
Public ModelSpec As String                                 'ModelM
Public SysVerSpec As String                                'SysVerM
Public FlashInfoSpec As String                             'FlashInfoM
Public HardwareVerSpec As String                           'HardwareVerM
Public DimensionSpec As String                             'DimensionM
Public ChannelSpec As String                               'ChannelM
Public PartitionVerSpec As String                          'PartitionVerM
Public TwoPointFourGVerSpec As String                      '24GVerM
Public PanelSpec As String                                 'PanelM
Public CarrierSpec As String                               'CarrierM
Public AreaSpec As String                                  'AreaM
Public HDCPSpec As String                                  'HDCPM
Public ResolutionSpec As String                            'ResolutionM
Public MACAddrSpec As String                               'MACAddrM
Public DeviceKeySpec As String                             'DeviceKeyM


Public IsStop As Boolean
Public IsACK As Boolean

Public strCurrentModelName As String
Public strDataVersion As String

Public SetTVCurrentComID As Integer
Public SetData As Integer
Public SetDay As Integer

Public IsSNWriteSuccess As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long



