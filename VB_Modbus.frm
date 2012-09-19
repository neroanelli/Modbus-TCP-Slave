VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form fmMain 
   Caption         =   "Modbus RTU/TCP协议 客户端-Mister.T"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VB_Modbus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   16335
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   5520
      TabIndex        =   38
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnOPC 
      Caption         =   "Connect"
      Height          =   495
      Left            =   5520
      TabIndex        =   37
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "OPCServer Interface"
      Height          =   2775
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Width           =   7455
      Begin VB.CheckBox DataChgChk 
         Caption         =   "启用订阅数据更新"
         Height          =   375
         Left            =   5400
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   1320
         TabIndex        =   39
         Top             =   1590
         Width           =   3495
      End
      Begin VB.TextBox txtOPCAddress 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "127.0.0.1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "OPCServer Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   33
         Top             =   1080
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "OPCServer Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   34
         Top             =   120
         Width           =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   5
         X1              =   -120
         X2              =   7440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   690
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OPC Name"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   1650
         Width           =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   4
         X1              =   0
         X2              =   7440
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.CommandButton btnTCP 
      Caption         =   "Listen"
      Height          =   495
      Left            =   14040
      TabIndex        =   30
      Top             =   2400
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComPort 
      Left            =   11400
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame FrameCom 
      Caption         =   "COM Interface"
      Height          =   2775
      Left            =   7680
      TabIndex        =   17
      Top             =   3120
      Width           =   8295
      Begin VB.CommandButton btnOpenPort 
         Caption         =   "Open Port"
         Height          =   495
         Left            =   6360
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox StopCb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox DataCb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox CheckCb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox RateCb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox PortCb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stop Bit:"
         Height          =   195
         Index           =   8
         Left            =   5595
         TabIndex        =   28
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Bit:"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check Way:"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   26
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Baud Rate:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Local Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   18
         Top             =   120
         Width           =   1350
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   3
         X1              =   0
         X2              =   8280
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COM Port:"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   690
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         X1              =   0
         X2              =   8280
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Data"
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   7455
      Begin VB.CommandButton btnReadExcel 
         Caption         =   "Read from Excel"
         Height          =   495
         Left            =   2160
         TabIndex        =   16
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton btnWriteExcel 
         Caption         =   "Write to Excel"
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   5640
         TabIndex        =   14
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox txtAutoData 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5175
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9128
         _Version        =   393216
      End
   End
   Begin VB.Frame FrameTCP 
      Caption         =   "TCP Interface"
      Height          =   2775
      Left            =   7680
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.TextBox txtIPPort 
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   9
         Text            =   "502"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtIPAddress 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Text            =   "127.0.0.1"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtIPPort 
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   4
         Text            =   "502"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtIPAddress 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "127.0.0.1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remote Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Local Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Top             =   120
         Width           =   1350
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   0
         X1              =   0
         X2              =   8280
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP Port"
         Height          =   210
         Index           =   1
         Left            =   4920
         TabIndex        =   8
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1650
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP Port"
         Height          =   210
         Index           =   0
         Left            =   4920
         TabIndex        =   3
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   690
         Width           =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   -120
         X2              =   8280
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11520
      Top             =   8640
   End
   Begin MSWinsockLib.Winsock Wsk_Server 
      Index           =   0
      Left            =   11520
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   7920
      TabIndex        =   40
      Top             =   6120
      Width           =   8055
   End
   Begin VB.Menu 菜单 
      Caption         =   "&File"
      Begin VB.Menu Import_Cfg 
         Caption         =   "&ImportConfig"
      End
      Begin VB.Menu ImportItems 
         Caption         =   "&ImportItems"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "&Config"
      Begin VB.Menu TCP_Param 
         Caption         =   "&Tcp"
      End
      Begin VB.Menu Rtu_Param 
         Caption         =   "&Rtu"
      End
   End
   Begin VB.Menu Winsock_Start 
      Caption         =   "TcpStart"
   End
   Begin VB.Menu COM_Start 
      Caption         =   "COMStart"
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents opcserver As OPCClass
Attribute opcserver.VB_VarHelpID = -1
Dim SubOPC As Boolean '启用订阅更新

Dim xlsFilePath As String
Dim InputS(), CoilS(), InputR(), HoldR(), rowLen, colLen
Dim ClientNum As Integer, lhostName As String
Dim i As Integer
Dim ComNumber, ComBit, ComBps, ComStopBit, ComCheck, RemoteIP, OPCServerIP, LocalPort, RemotePort As String
Dim CSVP As New CSVParse

'Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long  '判断数组为空


'----------【导入配置文件】----------
Private Sub ImportCfgFile()

 TagSum = CInt(ReadFile_INI("RealNum", "Count"))
 ComNumber = ReadFile_INI("Configuration", "ComNumber")
 ComBit = ReadFile_INI("Configuration", "ComBit")
 ComBps = ReadFile_INI("Configuration", "ComBps")
 ComStopBit = ReadFile_INI("Configuration", "ComStopBit")
 ComCheck = ReadFile_INI("Configuration", "ComCheck ")
 RemoteIP = ReadFile_INI("Configuration", "RemoteIP ")
 LocalPort = ReadFile_INI("Configuration", "LocalPort ")
 RemotePort = ReadFile_INI("Configuration", "RemotePort ")
 OPCServerIP = ReadFile_INI("Configuration", "OPCServerIP ")
 
 PortCb.ListIndex = SendMessage(PortCb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(ComNumber))
 CheckCb.ListIndex = SendMessage(CheckCb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(ComCheck))
 RateCb.ListIndex = SendMessage(RateCb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(ComBps))
 DataCb.ListIndex = SendMessage(DataCb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(ComBit))
 StopCb.ListIndex = SendMessage(StopCb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(ComStopBit))
     
     txtIPPort(0).Text = LocalPort
     txtIPAddress(1).Text = RemoteIP
     txtIPPort(1).Text = RemotePort

End Sub



Private Sub btnOPC_Click()
Set opcserver = New OPCClass
    If opcserver.Connect(Combo1.Text, txtOPCAddress.Text) Then
         
    Dim Tag As String
'Dim Savetime  As Double


   ' tag = LvListView.ListItems(1).SubItems(1)
   Tag = MSFlexGrid2.TextMatrix(1, 3)
    For i = 2 To 11
        Tag = Tag & "," & MSFlexGrid2.TextMatrix(i, 3)
    Next i
    opcserver.AddItem Tag, SubOPC
    
'Label1:
'opcserver.AsyncRead
'Dim Savetime As Double
'timeBeginPeriod 1
'Savetime = timeGetTime
'While timeGetTime < Savetime + 2000
''     If ss = True Then
''     timeEndPeriod 1
''     opcserver.Disconnect
''     Set opcserver = Nothing
''        Exit Sub
''    End If
'DoEvents
'Wend
'GoTo Label1


    'btnAddItem.Enabled = True
    End If

End Sub

Private Sub btnRefresh_Click()
On Error Resume Next
    Combo1.Clear
    Dim i As Integer
    Dim servername As Variant
    Dim opcs As OPCClass

Set opcs = New OPCClass
    servername = opcs.GetOPCServers(txtOPCAddress.Text)
Set opcs = Nothing
   For i = LBound(servername) To UBound(servername)
       Combo1.AddItem servername(i), i - 1
   Next i
End Sub



Private Sub DataChgChk_Click()
    If DataChgChk.Value = vbChecked Then
        tmUpdate.Enabled = False
        SubOPC = True
    Else
         tmUpdate.Enabled = True
        
        SubOPC = False
        
    End If
End Sub

Private Sub Import_Cfg_Click()
Call ImportCfgFile
End Sub

'----------【导入变量文件】----------
Private Sub ImportItems_Click()
ReDim Data(1 To TagSum) As TagData
 CSVP.FileName = App.Path + "\OPC_CONF.csv"
    CSVP.LoadNextLine
    For i = 1 To TagSum
        CSVP.LoadNextLine
        Data(i).Tag = CSVP.GetField(2)
        Data(i).TagName = CSVP.GetField(3)

'        Set itx = lvListView.ListItems.Add(, , CSVP.GetField(1))
'        itx.SubItems(1) = HourData(i).TagName
'        itx.SubItems(3) = HourData(i).HH
'        itx.SubItems(4) = HourData(i).LL
        MSFlexGrid2.TextMatrix(i, 3) = Data(i).Tag
    Next i

Set CSVP = Nothing
End Sub
'----------OPC数据更新
Private Sub opcserver_DataChange( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, ItemValues() As Variant, _
    Qualities() As Long, TimeStamps() As Date)
    Dim strBuf As String
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim nDrawHeight As Integer
    Dim sglScale As Single
    Dim i As Integer

    
        For i = 1 To UBound(ItemValues)

 MSFlexGrid2.TextMatrix(i, 4) = ItemValues(i)
         Next i

        

End Sub
'----------【退出系统】----------
Private Sub btnExit_Click()


    End
End Sub


Private Sub btnOpenPort_Click() '打开串口
Dim Settings As String
Dim j As Long
Dim SaveCfg As Boolean

On Error GoTo ErrP

    Me.Caption = "Modbus Slave--COM"
    If btnOpenPort.Caption = "Open Port" Then         ''打开串口
        Select Case CheckCb.ListIndex
        Case 0
            Settings = "N"
        Case 1
            Settings = "O"
        Case 2
            Settings = "E"
        End Select

        Settings = RateCb.Text + "," + Settings + "," + DataCb.Text + "," + StopCb.Text
        MSComPort.CommPort = PortCb.ListIndex + 1
        MSComPort.OutBufferSize = 1024
        MSComPort.InBufferSize = 1024
        MSComPort.InputMode = 1
        MSComPort.InputLen = 0
        MSComPort.InBufferCount = 0
        MSComPort.SThreshold = 1
        MSComPort.RThreshold = 1
        MSComPort.Settings = Settings
        MSComPort.PortOpen = True

        If MSComPort.PortOpen = True Then
            OpenFlag = True
            Import_Cfg.Enabled = False
            'ComTimer.Enabled = True
''            Option2(0).Enabled = False
''            Option2(1).Enabled = False
'            Frame1(1).Enabled = True
            btnOpenPort.Caption = "Close Port"
           ' CmdSend.Enabled = True
            Me.Caption = "Modbus(RTU) Tool--Slave  " + PortCb.Text + ":" + Settings
           SaveCfg = WriteFile_INI("Configuration", "ComNumber", PortCb.Text) And WriteFile_INI("Configuration", "ComBps", RateCb.Text) And WriteFile_INI("Configuration", "ComBit", DataCb.Text) And WriteFile_INI("Configuration", "ComStopBit", StopCb.Text) And WriteFile_INI("Configuration", "ComCheck", CheckCb.Text)
           
            If SaveCfg = False Then GoTo ErrLog
            
        End If
    Else
        btnOpenPort.Caption = "Open Port"             ''关闭串口
     
        'CmdSend.Enabled = False
        If OpenFlag = True Then
            OpenFlag = False
           If ConnectFlag = False Then Import_Cfg.Enabled = True
            MSComPort.PortOpen = False
        End If
'        Option2(0).Enabled = True
'        Option2(1).Enabled = True
'        Frame2.Enabled = False
      '  AutoTimer.Enabled = False
        'OverTimer.Enabled = False
       ' ComTimer.Enabled = False
'        Frame1(1).Enabled = False
'        Check2.Value = 0
    End If
    Exit Sub
ErrLog:
    
   WriteLog ("Err on Saving ConfigFile")
ErrP:
    MsgBox err.Description, vbCritical + vbOKOnly, "COM"
End Sub
Private Sub Initface()
Dim i As Integer

    For i = 1 To 15
        PortCb.AddItem "COM" + CStr(i)
    Next i
    PortCb.ListIndex = 0
    
    With RateCb             '波特率
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "57600"
        .AddItem "115200"
        .ListIndex = 6
    End With
    
    With CheckCb
        .AddItem "None"     '校验方式
        .AddItem "Odd"
        .AddItem "Even"
        .ListIndex = 0
    End With
        
    DataCb.AddItem "6"          '数据位
    DataCb.AddItem "7"
    DataCb.AddItem "8"
    DataCb.ListIndex = 2
    
'   stopcb.AddItem "1.5"       '停止位
    StopCb.AddItem "1"
    StopCb.AddItem "2"
    StopCb.ListIndex = 0
    
'    For i = 1 To 255
'        CbSlaveAddr.AddItem i
'    Next i
'    CbSlaveAddr.ListIndex = 0
'    SAddr = 1
End Sub

'----------【读取Excel数据】----------
Private Sub btnReadExcel_Click()
    Call xlsRead(xlsFilePath)
    Call initMsFlexGrid
End Sub

Private Sub btnTCP_Click()
If btnTCP.Caption = "Listening..." Then
Wsk_Server(0).Close
btnTCP.Caption = "Listen"
ConnectFlag = False
  If OpenFlag = False Then Import_Cfg.Enabled = True

Else
    txtIPAddress(0).Text = Wsk_Server(0).LocalIP
    Wsk_Server(0).LocalPort = txtIPPort(0).Text
    Wsk_Server(0).RemoteHost = txtIPAddress(1).Text
    Wsk_Server(0).RemotePort = txtIPPort(1).Text
    '程序启动时侦听
    Wsk_Server(0).Listen
    lhostName = Wsk_Server(0).LocalHostName & ":"
   btnTCP.Caption = "Listening..."
   ConnectFlag = True
    Import_Cfg.Enabled = False
    End If
    
End Sub

'----------【写入Excel数据】----------
Private Sub btnWriteExcel_Click()
    On Error Resume Next
    Dim i, j, m, N
    Dim xlApp, xlBook, xlSheet1
    Dim xlsFileOpen
    xlsFileOpen = 0
    Set xlApp = GetObject(, "Excel.Application")    '取得当前运行的Excel对象
    For i = 1 To xlApp.Workbooks.Count
        Set xlBook = xlApp.Workbooks(i)             '当前Excel打开的工作簿文件
        If err.number <> 0 Then Exit For
        Debug.Print xlBook.fullname
        If xlBook.fullname = xlsFilePath Then
            xlBook.save
            xlsFileOpen = 1
            Exit For
        End If
    Next
    If xlsFileOpen = 0 Then
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(xlsFilePath)
    End If
    Set xlSheet1 = xlBook.Worksheets(1)
    For j = 2 To colLen
        For i = 2 To rowLen
            xlSheet1.Cells(i, j).Value = MSFlexGrid2.TextMatrix(i - 1, j - 1)
        Next
    Next
    xlBook.save
    If xlsFileOpen = 0 Then xlApp.quit
    Set xlApp = Nothing '交还控制给Excel
End Sub
'----------Form Load and Initiation----------
Private Sub Form_Load()
    '-------------------------------------------------

    txtAutoData.Visible = False
    Initface
    xlsFilePath = App.Path & "\ModbusData.xls"
    Call xlsRead(xlsFilePath)
    Call initMsFlexGrid
    '-------------------------------------------------

    Timer1.Enabled = False
    Call ImportCfgFile
    
    
End Sub

'----------程序初始化，将Excel表格数据读入----------
Private Sub xlsRead(ByVal xlsFilePath As String)
    On Error Resume Next
    Dim i, j
    Dim xlApp, xlBook, xlSheet1
    Dim xlsFileOpen
    
    xlsFileOpen = 0
    Set xlApp = GetObject(, "Excel.Application")               '取得当前运行的Excel对象
    If err.number = 0 Then
        For i = 1 To xlApp.Workbooks.Count
            Set xlBook = xlApp.Workbooks(i)             '当前Excel打开的工作簿文件
            If err.number <> 0 Then Exit For
            Debug.Print xlBook.fullname
            If xlBook.fullname = xlsFilePath Then
                xlBook.save
                xlsFileOpen = 1
                Exit For
            End If
        Next
    End If
    If xlsFileOpen = 0 Then
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(xlsFilePath)
    End If
    Set xlSheet1 = xlBook.Worksheets(1)
    rowLen = xlSheet1.usedrange.Rows.Count
    colLen = xlSheet1.usedrange.Columns.Count
    ReDim InputS(rowLen)
    ReDim CoilS(rowLen)
    ReDim InputR(rowLen)
    ReDim HoldR(rowLen)
    For i = 1 To xlSheet1.usedrange.Rows.Count
        InputS(i - 1) = xlSheet1.Cells(i, 2).Value
        CoilS(i - 1) = xlSheet1.Cells(i, 3).Value
        InputR(i - 1) = xlSheet1.Cells(i, 4).Value
        HoldR(i - 1) = xlSheet1.Cells(i, 5).Value
    Next
    xlBook.save
    If xlsFileOpen = 0 Then xlApp.quit
    Set xlApp = Nothing '交还控制给Excel
End Sub

'----------程序初始化，将Excel表格数据写入控件----------
Private Sub initMsFlexGrid()
    Dim i, j
    '-------------------初始化全自动运行表头----------------------------
    MSFlexGrid2.Rows = rowLen  '设置MSFlexGrid 表格的总行数
    MSFlexGrid2.Cols = colLen  '设置MSFlexGrid 表格的总列数
    '设置MSFlexGrid 表格的列宽
    For i = 0 To rowLen - 1
        MSFlexGrid2.RowHeight(i) = 300
    Next
    MSFlexGrid2.ColWidth(0) = 850
    For i = 1 To colLen - 1
        MSFlexGrid2.ColWidth(i) = 1500
    Next
    '设置MSFlexGrid 表格的固定行数
    MSFlexGrid2.FixedRows = 1
    '设置MSFlexGrid 表格的固定列数
    MSFlexGrid2.FixedCols = 1
    '设置MSFlexGrid 表格的表头信息
    MSFlexGrid2.TextMatrix(0, 0) = "No."
    MSFlexGrid2.TextMatrix(0, 1) = InputS(0)
    MSFlexGrid2.TextMatrix(0, 2) = CoilS(0)
    MSFlexGrid2.TextMatrix(0, 3) = InputR(0)
    MSFlexGrid2.TextMatrix(0, 4) = HoldR(0)
    
    '为MSFlexGrid 表格设置序号,并读入数据
    For i = 1 To rowLen - 1
        MSFlexGrid2.TextMatrix(i, 0) = i
        MSFlexGrid2.TextMatrix(i, 1) = InputS(i)
        MSFlexGrid2.TextMatrix(i, 2) = CoilS(i)
        MSFlexGrid2.TextMatrix(i, 3) = InputR(i)
        
        MSFlexGrid2.TextMatrix(i, 4) = HoldR(i)
    Next i
        '-------------------------------------------------------------------
End Sub





Private Sub Timer1_Timer()
For i = 1 To 800
MSFlexGrid2.TextMatrix(i, 4) = Rnd * 10000
Next
End Sub







'----------TCP通讯-客户端断开连接时，关闭连接----------
Private Sub Wsk_Server_Close(Index As Integer)
    On Error Resume Next
    Dim strWelc
    strWelc = "欢迎您的再次光临，再见！"
    Wsk_Server(ClientNum).SendData strWelc
    Wsk_Server(Index).Close
End Sub

'=============================================================================
'e.State属性
'   返回WinSock控件当前的状态
'
'   常数                    值      描述
'   sckClosed               0       缺省值,关闭。
'   SckOpen                 1       打开。
'   SckListening            2       侦听
'   sckConnectionPending    3       连接挂起
'   sckResolvingHost        4       识别主机。
'   sckHostResolved         5       已识别主机
'   sckConnecting           6       正在连接。
'   sckConnected            7       已连接。
'   sckClosing              8       同级人员正在关闭连接。
'   sckError                9       错误
'=============================================================================
'----------TCP通讯-接收客户端连接请求----------
Private Sub Wsk_Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    Dim i, j, strWelc
    strWelc = "欢迎您的光临！！"
    j = 0
    For i = 1 To ClientNum
        If Wsk_Server(i).State <> 7 Then
            Wsk_Server(i).Close
            j = i
        End If
    Next
    If j > 0 Then
        Wsk_Server(j).Accept requestID          '接受客户端的请求
        Wsk_Server(j).SendData strWelc
    Else
        ClientNum = ClientNum + 1
        Load Wsk_Server(ClientNum)              '载入一个新的socket控件
        Wsk_Server(ClientNum).Accept requestID  '接受客户端的请求
        Wsk_Server(ClientNum).SendData lhostName & strWelc
    End If
End Sub

'----------TCP通讯-接收客户端数据----------
Private Sub Wsk_Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim i, j As Integer
    Dim s As String
    Dim s1 As String
    Dim IO As Byte
    Dim strdata() As Byte
    i = Wsk_Server(Index).BytesReceived
    ReDim strdata(i)
    Wsk_Server(Index).GetData strdata, vbByte, i
    For j = 0 To i - 1
        s = s + " " + Right("000" & strdata(j), 3)
        s1 = s1 & " " & strdata(j)
    Next
    Debug.Print "server index-" & Index & " : " & s
    Call Ack_Server(strdata(), Index)
End Sub
'----------Modbus通讯-接收客户端数据----------
Private Sub MSComPort_OnComm()
Dim bytInput() As Byte
  Dim intInputLen As Integer
  Dim N As Integer
  Dim teststring As String
  Dim s1 As String
  Dim AscFlag As Boolean
  Dim mo As Boolean
  Dim t1, t2, t3, t4, t5 As String
  AscFlag = True
  mo = True
  
  Select Case MSComPort.CommEvent
    Case comEvReceive
      If mo = True Then
        MSComPort.InputMode = 1                    '0：文本方式，1：二进制方式
      Else
        MSComPort.InputMode = 0                    '0：文本方式，1：二进制方式
      End If
            
      intInputLen = MSComPort.InBufferCount
      bytInput = MSComPort.Input
            
      If AscFlag = True Then
        For N = 0 To intInputLen - 1
          s1 = s1 & " " & IIf(Len(Hex$(bytInput(N))) > 1, Hex$(bytInput(N)), "0" & Hex$(bytInput(N)))
        Next N
        t1 = Crc_16(MidB(bytInput, 1, 6))
        t4 = t1 \ 256
        t5 = t1 Mod 256
        t2 = bytInput(6)
        t3 = bytInput(7)
      Else
        teststring = bytInput
        s1 = s1 + teststring
        
      End If
     Debug.Print "ComData" & " : " & s1
   
      Call Ack_Server_RTU(bytInput())
  End Select


End Sub

'----------TCP通讯-客户端出现通讯故障----------
Private Sub Wsk_Server_Error(Index As Integer, ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    Wsk_Server(Index).Close

End Sub

'---------------- acknowledge to server(TCP) ------------------------------
Private Sub Ack_Server(strdata() As Byte, Index As Integer)
    '----------------response to server----------
    On Error GoTo ErrProc
    Dim AckStr() As Byte
    Dim i, j, stAddr, ackLen, a1, a2, a3, a4 As Integer
    Dim TmpStr As String
    Dim TmpStr1 As String
    Dim a As Single
    Dim FunctionCode As Integer
    Dim AckLength As Integer
    FunctionCode = strdata(7)
    AckLength = strdata(11)
        Select Case FunctionCode
        Case 1, 2:          ackLen = 8 + AckLength \ 8 + 1   '读线圈
        Case 3, 4:          ackLen = 9 + AckLength * 2      '读寄存器
        Case 5, 6, 15, 16:  ackLen = 12                      '写寄存器
        
    End Select

    ReDim AckStr(ackLen) As Byte
    AckStr(0) = strdata(0)      '交换识别号高字节，通常为 0
    AckStr(1) = strdata(1)      '交换识别号低字节，通常为 0
    '-------------------------------------
    AckStr(2) = 0               '协议识别号高字节，为 0
    AckStr(3) = 0               '协议识别号低字节，为 0
    AckStr(4) = 0               '字节长度高字节
    '-------------------------------------
    Dim strSend As String
    
    If FunctionCode = 1 Or FunctionCode = 2 Then        '=========== 线圈
        AckStr(5) = strdata(11) \ 8 + 1 + 3
        AckStr(6) = strdata(6)
        AckStr(7) = strdata(7)
        AckStr(8) = strdata(11) \ 8 + 1
        Dim temp As Byte, N As Integer
        For i = 9 To 9 + strdata(11) \ 8
            temp = 0: N = 0
            For j = strdata(8) * 256 + strdata(9) + (i - 9) * 8 + 1 To strdata(8) * 256 + strdata(9) + (i - 9) * 8 + 8
                If j > strdata(8) * 256 + strdata(9) + strdata(10) * 256 + strdata(11) Then Exit For
                Select Case strdata(7)
                    Case 1: temp = temp + MSFlexGrid2.TextMatrix(j, 1) * 2 ^ N
                    Case 2: temp = temp + MSFlexGrid2.TextMatrix(j, 2) * 2 ^ N
                End Select
                N = N + 1
            Next
            AckStr(i) = temp
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 3 Or FunctionCode = 4 Then               '=========== 寄存器
        AckStr(5) = strdata(11) * 2 + 3    '以下字节长度低字节
        AckStr(6) = strdata(6)             '单元识别号，缺省为 255
        AckStr(7) = strdata(7)             '读多个寄存器命令代码
        AckStr(8) = strdata(11) * 2        '读数据的起始地址高字节
        stAddr = strdata(9) + strdata(8) * 16
        j = 1
        For i = 9 To ackLen - 1 Step 4
            Select Case strdata(7)
                Case 3:   '返回数据
'                    AckStr(i) = Right("00" & (MSFlexGrid2.TextMatrix(stAddr + j, 4) \ 256), 3)         '数据高字节
'                    AckStr(i + 1) = Right("000" & MSFlexGrid2.TextMatrix(stAddr + j, 4) Mod 256, 3)    '数据低字节
                     'TmpStr = OTC2Single(MSFlexGrid2.TextMatrix(stAddr + j, 4))
                     a = Val(MSFlexGrid2.TextMatrix(stAddr + j, 4))
                     CopyMemory AckStr(i), a, 4
                     a1 = AckStr(i)
                     a2 = AckStr(i + 1)
                     a3 = AckStr(i + 2)
                     a4 = AckStr(i + 3)
                     AckStr(i) = a2
                     AckStr(i + 1) = a1
                     AckStr(i + 2) = a4
                     AckStr(i + 3) = a3
                    ' AckStr(i) = OTC2Single()         '数据高字节
                   '  AckStr(i + 1) = Right("000" & TmpStr, 2)   '数据低字节
                    TmpStr = Hex(AckStr(9)) & " " & Hex(AckStr(10)) & " " & Hex(AckStr(11)) & " " & Hex(AckStr(12))
                    TmpStr1 = OTC2Single(MSFlexGrid2.TextMatrix(stAddr + j, 4))
                    Debug.Print TmpStr
                    Debug.Print TmpStr1
                Case 4:
                    AckStr(i) = Right("00" & (MSFlexGrid2.TextMatrix(stAddr + j, 3) \ 256), 3)         '数据高字节
                    AckStr(i + 1) = Right("000" & MSFlexGrid2.TextMatrix(stAddr + j, 3) Mod 256, 3)    '数据低字节
            End Select
            j = j + 1
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 5 Or FunctionCode = 6 Then          '===========
        For i = 5 To 11
            AckStr(i) = strdata(i)
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 15 Or FunctionCode = 16 Then        '===========
        AckStr(5) = 6
        For i = 6 To 11
            AckStr(i) = strdata(i)
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    
    End If
    Debug.Print "sending data:" & vbTab & strSend
  
  
 
  Wsk_Server(Index).SendData AckStr
     
     
 
    '----------- update the form data ------------------
    If FunctionCode = 5 Then
        If strdata(10) = 255 Then strdata(10) = 1
        MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + 1, 1) = strdata(10)
    ElseIf FunctionCode = 6 Then
        MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + 1, 4) = strdata(10) * 256 + strdata(11)
    ElseIf FunctionCode = 15 Then
        Dim Temp1 As String
        N = N + 1
        For i = 13 To 13 + strdata(12) - 1
            Temp1 = Byte_to_BIN(strdata(i)) & Temp1
        Next
        j = 1: Temp1 = StrReverse(Temp1)
        Debug.Print Temp1
        For i = strdata(8) * 256 + strdata(9) To strdata(8) * 256 + strdata(9) + strdata(10) * 256 + strdata(11) - 1
            MSFlexGrid2.TextMatrix(i + 1, 1) = Left$(Temp1, 1)
            Temp1 = Mid$(Temp1, 2)
        Next
    ElseIf FunctionCode = 16 Then
        j = 1
        For i = 0 To strdata(12) - 1 Step 2
            MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + j, 4) = strdata(13 + i) * 256 + strdata(14 + i)
            j = j + 1
        Next
    End If
    Exit Sub
ErrProc:
    Debug.Print "传输数据失败:" & vbTab & err.Description
End Sub
'---------------- acknowledge to server(RTU) ------------------------------
Private Sub Ack_Server_RTU(strdata() As Byte)
    '----------------response to server----------
    On Error GoTo ErrProc
    Dim AckStr() As Byte
    Dim i, j, stAddr, ackLen, a1, a2, a3, a4 As Integer
    Dim TmpStr, CRCStr As String
    Dim TmpStr1 As String
    Dim a As Single
    Dim FunctionCode As Integer
    Dim AckLength As Integer
    FunctionCode = strdata(1)
    AckLength = strdata(5)
    
        Select Case FunctionCode
        Case 1, 2:          ackLen = 4 + AckLength \ 8 + 1   '读线圈
        Case 3, 4:          ackLen = 5 + AckLength * 2      '读寄存器
        Case 5, 6, 15, 16:  ackLen = 8                      '写寄存器
        
    End Select

    ReDim AckStr(ackLen) As Byte
    AckStr(0) = strdata(0)      '交换识别号高字节，通常为 0
    AckStr(1) = strdata(1)      '交换识别号低字节，通常为 0
    '-------------------------------------

    Dim strSend As String
    
    If FunctionCode = 1 Or FunctionCode = 2 Then        '=========== 线圈
        AckStr(2) = strdata(6) \ 8 + 1
        Dim temp As Byte, N As Integer
        For i = 3 To 3 + strdata(5) \ 8
            temp = 0: N = 0
            For j = strdata(2) * 256 + strdata(3) + (i - 3) * 8 + 1 To strdata(2) * 256 + strdata(3) + (i - 3) * 8 + 8
                If j > strdata(2) * 256 + strdata(3) + strdata(4) * 256 + strdata(5) Then Exit For
                Select Case FunctionCode
                    Case 1: temp = temp + MSFlexGrid2.TextMatrix(j, 1) * 2 ^ N
                    Case 2: temp = temp + MSFlexGrid2.TextMatrix(j, 2) * 2 ^ N
                End Select
                N = N + 1
            Next
            AckStr(i) = temp
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 3 Or FunctionCode = 4 Then               '=========== 寄存器
        AckStr(2) = strdata(5) * 2        '读数据的起始地址高字节
        stAddr = strdata(3) + strdata(2) * 16
        j = 1
        For i = 3 To ackLen - 5 Step 4
            Select Case FunctionCode
                Case 3:   '返回数据
'                    AckStr(i) = Right("00" & (MSFlexGrid2.TextMatrix(stAddr + j, 4) \ 256), 3)         '数据高字节
'                    AckStr(i + 1) = Right("000" & MSFlexGrid2.TextMatrix(stAddr + j, 4) Mod 256, 3)    '数据低字节
                     'TmpStr = OTC2Single(MSFlexGrid2.TextMatrix(stAddr + j, 4))
                     a = Val(MSFlexGrid2.TextMatrix(stAddr + j, 4))
                     CopyMemory AckStr(i), a, 4
                     a1 = AckStr(i)
                     a2 = AckStr(i + 1)
                     a3 = AckStr(i + 2)
                     a4 = AckStr(i + 3)
                     AckStr(i) = a2
                     AckStr(i + 1) = a1
                     AckStr(i + 2) = a4
                     AckStr(i + 3) = a3
                    ' AckStr(i) = OTC2Single()         '数据高字节
                   '  AckStr(i + 1) = Right("000" & TmpStr, 2)   '数据低字节
'                    TmpStr = Hex(AckStr(3)) & " " & Hex(AckStr(4)) & " " & Hex(AckStr(5)) & " " & Hex(AckStr(6))
'                    TmpStr1 = OTC2Single(MSFlexGrid2.TextMatrix(stAddr + j, 4))
'                    Debug.Print TmpStr
'                    Debug.Print TmpStr1
                Case 4:
                    AckStr(i) = Right("00" & (MSFlexGrid2.TextMatrix(stAddr + j, 3) \ 256), 3)         '数据高字节
                    AckStr(i + 1) = Right("000" & MSFlexGrid2.TextMatrix(stAddr + j, 3) Mod 256, 3)    '数据低字节
            End Select
            j = j + 1
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 5 Or FunctionCode = 6 Then          '===========
        For i = 2 To 6
            AckStr(i) = strdata(i)
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    ElseIf FunctionCode = 15 Or FunctionCode = 16 Then        '===========
        For i = 2 To 6
            AckStr(i) = strdata(i)
        Next
        For i = 0 To ackLen - 1
            strSend = strSend & " " & Right("000" & AckStr(i), 3)
        Next
    
    End If
      
        CRCStr = Crc_16(MidB(AckStr, 1, ackLen - 2))

      AckStr(ackLen - 2) = CRCStr Mod 256
      AckStr(ackLen - 1) = CRCStr \ 256
      Debug.Print "ComPort Sending data:" & vbTab & strSend
      MSComPort.Output = AckStr

  
    
    
  
    '----------- update the form data ------------------
    If FunctionCode = 5 Then
        If FunctionCode = 255 Then strdata(10) = 1
        MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + 1, 1) = strdata(10)
    ElseIf FunctionCode = 6 Then
        MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + 1, 4) = strdata(10) * 256 + strdata(11)
    ElseIf FunctionCode = 15 Then
        Dim Temp1 As String
        N = N + 1
        For i = 13 To 13 + strdata(12) - 1
            Temp1 = Byte_to_BIN(strdata(i)) & Temp1
        Next
        j = 1: Temp1 = StrReverse(Temp1)
        Debug.Print Temp1
        For i = strdata(8) * 256 + strdata(9) To strdata(8) * 256 + strdata(9) + strdata(10) * 256 + strdata(11) - 1
            MSFlexGrid2.TextMatrix(i + 1, 1) = Left$(Temp1, 1)
            Temp1 = Mid$(Temp1, 2)
        Next
    ElseIf FunctionCode = 16 Then
        j = 1
        For i = 0 To strdata(12) - 1 Step 2
            MSFlexGrid2.TextMatrix(strdata(8) * 256 + strdata(9) + j, 4) = strdata(13 + i) * 256 + strdata(14 + i)
            j = j + 1
        Next
    End If
    Exit Sub
ErrProc:
    Debug.Print "传输数据失败:" & vbTab & err.Description
End Sub

'----------【单击单元格】----------
Private Sub MSFlexGrid2_Click() '单击单元格
    '指定text1 控件在MSFlexGrid1 表格中的大小及位置
    txtAutoData.Width = MSFlexGrid2.CellWidth
    txtAutoData.Height = MSFlexGrid2.CellHeight
    txtAutoData.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
    txtAutoData.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
    '赋值给MSFlexGrid2.Text
    txtAutoData.Text = MSFlexGrid2.Text
    txtAutoData.SelStart = 0
    txtAutoData.SelLength = Len(txtAutoData.Text)
    'text1 可见
    txtAutoData.Visible = True
    'text1 获得焦点
    txtAutoData.SetFocus
End Sub

'----------【单元格回车】----------
Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
    MSFlexGrid2_Click
End Sub

'---------- keypress on textbox -----------
Private Sub txtAutoData_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyEscape Then '当按下ESC 键时
        txtAutoData.Visible = False 'text1 不可见
        MSFlexGrid2.SetFocus 'MSFlexGrid1 获得焦点
        Exit Sub
        End If
        If KeyAscii = vbKeyReturn Then '当按下回车键时
            '赋值给txtAutodata.text
            MSFlexGrid2.Text = txtAutoData.Text
            If MSFlexGrid2.Row < MSFlexGrid2.Rows Then
                MSFlexGrid2.Row = MSFlexGrid2.Row + 1
            Else
                MSFlexGrid2.Col = MSFlexGrid2.Col + 1
                MSFlexGrid2.Row = 1
        End If
        '指定text1 控件在MSFlexGrid2 表格中的大小及位置
        txtAutoData.Width = MSFlexGrid2.CellWidth
        txtAutoData.Height = MSFlexGrid2.CellHeight
        txtAutoData.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
        txtAutoData.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
        '赋值给MSFlexGrid2.Text
        txtAutoData.Text = MSFlexGrid2.Text
        txtAutoData.SelStart = 0
        txtAutoData.SelLength = Len(txtAutoData.Text)
        txtAutoData.SetFocus 'text1 获得焦点
    End If
End Sub

'---------- double click on the textbox then hide it ----------
Private Sub txtAutoData_DblClick()
    txtAutoData.Visible = False
End Sub

' 用途：将十进制转化为二进制
' 输入：Byte1(十进制数)
' 输入数据类型：byte
' 输出：Byte_to_BIN(二进制数)
' 输出数据类型：String
' 输入的最大数为255,输出最大数为1111 1111 (8个1)
Private Function Byte_to_BIN(Byte1 As Byte) As String
    Byte_to_BIN = ""
    Dim i As Integer
    For i = 0 To 7
        Byte_to_BIN = Byte1 Mod 2 & Byte_to_BIN
        Byte1 = Byte1 \ 2
    Next
End Function
Private Function OTC2Single(text1 As Single) As Variant

Dim hexData As String
  Dim i As Integer
  Dim a As Single
  Dim Buffer(3) As Byte
  a = Val(text1)
  CopyMemory Buffer(0), a, 4
  For i = 0 To 3
        If Len(Hex(Buffer(i))) = 1 Then
            hexData = "0" & Hex(Buffer(i)) + hexData
        Else
            hexData = Hex(Buffer(i)) + hexData
        End If
    Next
    OTC2Single = hexData
End Function
