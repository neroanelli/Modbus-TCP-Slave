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

Private WithEvents objserver As opcserver     '����OPCSERVER
Attribute objserver.VB_VarHelpID = -1
Private objGroups As OPCGroups                '����OPC��
Private WithEvents objtestgrp As OPCGroup     'ʹ�õ�OPC��
Attribute objtestgrp.VB_VarHelpID = -1
Private objItems As OPCItems                  'OPC��
Private LServerHandles() As Long             '�������˷��ص���Ŀ��������ڷ������˶�д����
Private lTransID_Rd As Long                  '�����첽��ȡ����ʱ������ɵ����ݷ��ʣ���Ӧ�ó�����
Private lCancelID_Rd As Long                 '����˷��е�����ȡ�����ʵı�ʶ��
Private lTransID_Wt As Long                  'ʶ����ɵ����ݷ���
Private lCancelID_Wt As Long                 '����ȡ�����ڷ����е�����
Private Num_All As Integer                   '�������Ŀ����
'Private lServerState As Boolean              '����״̬
Public Event AsyncReadComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date, Errors() As Long)
Public Event DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
Public Event AsyncWriteComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, Errors() As Long)


Public Function Connect(ByVal strProgID As String, Optional strNode As String) As Boolean
    On Error GoTo err
    If objserver Is Nothing Then
        ' ����һ��OPC����������
        Set objserver = New opcserver
    End If
    
    If objserver.ServerState = OPCDisconnected Then
        ' ����OPC������
       ' objserver.ServerState = opcconnected
        objserver.Connect strProgID, strNode
    End If
    
    If objGroups Is Nothing Then
        ' ����һ��OPC�鼯��
        Set objGroups = objserver.OPCGroups
    End If
    
    If objtestgrp Is Nothing Then
        ' ����һ��OPC��
        Set objtestgrp = objGroups.Add("TestGrp")
    End If
    Connect = True
    Exit Function
err:
Connect = False
MsgBox "�޷����ӵ�ָ��������", vbOKOnly
Disconnect
End Function
Public Sub Disconnect()
    On Error Resume Next
    ReDim lErrors(Num_All) As Long
    
    If Not objItems Is Nothing Then
        If objItems.Count > 0 Then
            ' ���OPC��
            objItems.Remove Num_All, LServerHandles, lErrors
        End If
        Set objItems = Nothing
    End If
    
    If Not objtestgrp Is Nothing Then
        ' ���OPC��
        objGroups.Remove "TestGrp"
        objGroups.RemoveAll
        Set objtestgrp = Nothing
    End If
    
    If Not objGroups Is Nothing Then
        Set objGroups = Nothing
    End If
    
    If Not objserver Is Nothing Then
        If objserver.ServerState <> OPCDisconnected Then
            ' �Ͽ�OPC������.
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
    
    ' ������״̬
    If IsA Then
        objtestgrp.IsActive = True
    Else
        objtestgrp.IsActive = False
    End If
    ' �������ͬ��֪ͨ
    objtestgrp.IsSubscribed = True
    
    ' ����OPC���
    Set objItems = objtestgrp.OPCItems
    
    
    
    ' �������ʶ��
    For i = 1 To Num_All
        strItemIDs(i) = ItemID(i - 1)
        lClientHandles(i) = i
    Next i
    ' ����OPC��
    Call objItems.AddItems(Num_All, strItemIDs, _
        lClientHandles, LServerHandles, lErrors)
        fmMain.Label3 = "���سɹ���"
        Exit Sub
err:
MsgBox "�������ӵ�ָ������", vbOKOnly
Disconnect
End Sub
Rem �����ݷ�����Ŀ���е�ǰ�棬����ǰ������ȡ���ٸ�
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
        ' ��ͬ�ڶ�ȡ
        lTransID_Rd = lTransID_Rd + 1

        For i = 1 To item_num
            ServerHandles(i) = LServerHandles(i)
        Next i

        objtestgrp.AsyncRead item_num, ServerHandles, _
            lErrors, lTransID_Rd, lCancelID_Rd
    End If

End Sub
Rem д���ݷ�����Ŀ���еĺ��棬���Ӻ�����д����ٸ���д��˳��Ϊ����
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
        
        ' ��ͬ��д��
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