VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OPCClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private WithEvents objserver As opcserver     '定义OPCSERVER
Attribute objserver.VB_VarHelpID = -1
Private objGroups As OPCGroups                '定义OPC组
Private WithEvents objtestgrp As OPCGroup     '使用的OPC组
Attribute objtestgrp.VB_VarHelpID = -1
Private objItems As OPCItems                  'OPC项
Private LServerHandles() As Long             '服务器端返回的项目句柄，用于服务器端读写数据
Private lTransID_Rd As Long                  '用于异步读取数据时区分完成的数据访问，由应用程序发行
Private lCancelID_Rd As Long                 '服务端发行的用于取消访问的标识符
Private lTransID_Wt As Long                  '识别完成的数据访问
Private lCancelID_Wt As Long                 '用于取消正在访问中的数据
Private Num_All As Integer                   '加入的项目总数
'Private lServerState As Boolean              '连接状态
Public Event AsyncReadComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date, Errors() As Long)
Public Event DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
Public Event AsyncWriteComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, Errors() As Long)


Public Function Connect(ByVal strProgID As String, Optional strNode As String) As Boolean
    On Error GoTo err
    If objserver Is Nothing Then
        ' 建立一个OPC服务器对象
        Set objserver = New opcserver
    End If
    
    If objserver.ServerState = OPCDisconnected Then
        ' 连接OPC服务器
       ' objserver.ServerState = opcconnected
        objserver.Connect strProgID, strNode
    End If
    
    If objGroups Is Nothing Then
        ' 建立一个OPC组集合
        Set objGroups = objserver.OPCGroups
    End If
    
    If objtestgrp Is Nothing Then
        ' 添加一个OPC组
        Set objtestgrp = objGroups.Add("TestGrp")
    End If
    Connect = True
    Exit Function
err:
Connect = False
MsgBox "无法连接到指定服务器", vbOKOnly
Disconnect
End Function
Public Sub Disconnect()
    On Error Resume Next
    ReDim lErrors(Num_All) As Long
    
    If Not objItems Is Nothing Then
        If objItems.Count > 0 Then
            ' 清除OPC项
            objItems.Remove Num_All, LServerHandles, lErrors
        End If
        Set objItems = Nothing
    End If
    
    If Not objtestgrp Is Nothing Then
        ' 清除OPC组
        objGroups.Remove "TestGrp"
        objGroups.RemoveAll
        Set objtestgrp = Nothing
    End If
    
    If Not objGroups Is Nothing Then
        Set objGroups = Nothing
    End If
    
    If Not objserver Is Nothing Then
        If objserver.ServerState <> OPCDisconnected Then
            ' 断开OPC服务器.
            objserver.Disconnect
        End If
        
        Set objserver = Nothing
    End If
        
End Sub

Public Sub AddItem(ByVal items As String, ByVal IsA As Boolean)
    On Error GoTo err
    Dim ItemID() As String
    Dim i As Integer
    ItemID() = Split(items, ",")
    Num_All = UBound(ItemID) - LBound(ItemID) + 1
    ReDim strItemIDs(Num_All) As String
    ReDim lClientHandles(Num_All) As Long
    ReDim lErrors(Num_All) As Long
    
 
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If Not objItems Is Nothing Then
        If objItems.Count > 0 Then
            Exit Sub
        End If
    End If
    
    ' 设置组活动状态
    If IsA Then
        objtestgrp.IsActive = True
    Else
        objtestgrp.IsActive = False
    End If
    ' 启动组非同期通知
    objtestgrp.IsSubscribed = True
    
    ' 建立OPC项集合
    Set objItems = objtestgrp.OPCItems
    
    
    
    ' 生成项标识符
    For i = 1 To Num_All
        strItemIDs(i) = ItemID(i - 1)
        lClientHandles(i) = i
    Next i
    ' 添加OPC项
    Call objItems.AddItems(Num_All, strItemIDs, _
        lClientHandles, LServerHandles, lErrors)
        fmMain.Label3 = "加载成功！"
        Exit Sub
err:
MsgBox "不能连接到指定的项", vbOKOnly
Disconnect
End Sub
Rem 读数据放在项目序列的前面，即从前面数读取多少个
Public Sub AsyncRead() '(ByVal item_num As Long)
    
'    If item_num > Num_All Then
'       item_num = Num_All
'    End If
    Dim item_num As Long
    item_num = Num_All
    ReDim lErrors(item_num) As Long
    ReDim ServerHandles(item_num) As Long
    Dim i As Integer

    If objtestgrp Is Nothing Then
        Exit Sub
    End If

    If objtestgrp.OPCItems.Count > 0 Then
        ' 非同期读取
        lTransID_Rd = lTransID_Rd + 1

        For i = 1 To item_num
            ServerHandles(i) = LServerHandles(i)
        Next i

        objtestgrp.AsyncRead item_num, ServerHandles, _
            lErrors, lTransID_Rd, lCancelID_Rd
    End If

End Sub
Rem 写数据放在项目序列的后面，即从后面数写入多少个，写入顺序为正序
Public Sub AsyncWrite(ByRef vtItemValues() As Variant)
    Dim item_num As Integer
    item_num = UBound(vtItemValues) - LBound(vtItemValues) + 1
    If item_num > Num_All Then
       item_num = Num_All
    End If
    
    ReDim lHandle(item_num) As Long
    Dim i As Integer
    ReDim lErrors(item_num) As Long
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If objtestgrp.OPCItems.Count > 0 Then
        For i = 1 To item_num
            lHandle(i) = LServerHandles(i + Num_All - item_num)
        Next i
        
        ' 非同期写入
        lTransID_Wt = lTransID_Wt + 1
        objtestgrp.AsyncWrite item_num, lHandle(), vtItemValues, _
                lErrors, lTransID_Wt, lCancelID_Wt
    End If

End Sub
Public Function GetOPCServers(Optional Node As String) As Variant
 On Error GoTo lal
       Dim i As Integer
       Dim opcs As opcserver
       Set opcs = New opcserver
       GetOPCServers = opcs.GetOPCServers(Node)
       Set opcs = Nothing
       Exit Function
lal:
       fmMain.Label3 = Now & "--" & err.Description
End Function



Private Sub objtestgrp_AsyncReadComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date, Errors() As Long)
RaiseEvent AsyncReadComplete(TransactionID, NumItems, ClientHandles(), ItemValues(), Qualities(), TimeStamps(), Errors())
End Sub

Private Sub objtestgrp_AsyncWriteComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, Errors() As Long)
RaiseEvent AsyncWriteComplete(TransactionID, NumItems, ClientHandles(), Errors())
End Sub
Private Sub objtestgrp_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
RaiseEvent DataChange(TransactionID, NumItems, ClientHandles(), ItemValues(), Qualities(), TimeStamps())
End Sub
