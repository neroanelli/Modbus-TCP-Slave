VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Section1]
'Key1 = Value1
'Key2 = Value2
'Key3 = Vlaue3
'
'[Section2]
'Key1 = Value1
'Key2 = Value5
'Key4 = Value4
'Key5=...
'
'...

'Private Declare Function GetPrivateProfileInt Lib "kernel32" _
'Alias "GetPrivateProfileIntA" ( _       ' ��������ȡ�ĳ�����ֵ
'    ByVal lpApplicationName As String, _    ' Ҫ��ȡ�Ķ� (Section) ����
'    ByVal lpKeyName As String, _            ' Ҫ��ȡ�ĵļ� (Key) ����
'    ByVal nDefault As Long, _               ' ָ��Ĭ��ֵ�������ȡʱ�������򷵻ظ�ֵ
'    ByVal lpFileName As String) As Long     ' ָ��Ҫ���� INI �ļ���
'
'Private Declare Function GetPrivateProfileString Lib "kernel32" _
'Alias "GetPrivateProfileStringA" ( _    ' ��������ȡ���ַ���ֵ����ʵ����
'    ByVal lpApplicationName As String, _    ' Ҫ��ȡ�Ķ� (Section) ����
'    ByVal lpKeyName As Any, _               ' Ҫ��ȡ�ĵļ� (Key) ����
'    ByVal lpDefault As String, _            ' ָ��Ĭ��ֵ�������ȡʱ�������򷵻ظ�ֵ
'    ByVal lpReturnedString As String, _     ' ָ�����շ���ֵ���ַ�������
'    ByVal nSize As Long, _                  ' ָ�������ַ���ֵ����󳤶�
'    ByVal lpFileName As String) As Long     ' ָ��Ҫ���� INI �ļ���
'
'Private Declare Function WritePrivateProfileString Lib "kernel32" _
'Alias "WritePrivateProfileStringA" ( _  ' ����ɹ����ط� 0 ֵ��ʧ�ܷ��� 0
'    ByVal lpApplicationName As String, _    ' Ҫд��Ķ� (Section) ����
'    ByVal lpKeyName As Any, _               ' Ҫд��ĵļ� (Key) ����
'    ByVal lpString As Any, _                ' Ҫд���ֵ (Value)�����ַ�����ʾ
'    ByVal lpFileName As String) As Long     ' ָ��Ҫд�� INI �ļ���
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private IniFileName As String    '�ļ���
Public ErrorMsg As String        '������Ϣ




Private Sub Class_Initialize()
    IniFileName = vbNullString
    ErrorMsg = vbNullString

End Sub

Public Sub SpecifyIni(FilePathName As String)      'ָ�� INI �ļ����� CIniFile
    IniFileName = Trim(FilePathName)
End Sub

 Private Function NoIniFile() As Boolean  '�ж��Ƿ��Ѿ�ָ���� INI �ļ���
    NoIniFile = True
    If IniFileName = vbNullString Then
        ErrorMsg = "û��ָ�� INI �ļ�"
        Exit Function
    End If
    ErrorMsg = vbNullString
    NoIniFile = False
End Function

'�÷����� INI �ļ���д��һ����ֵ���ɹ����� True��ʧ�ܷ��� False��
'���� WritePrivateProfileString ����Ҫ�������ļ�����һ���������ṩ֮�⣬
'��Ҫ�ṩ������������ֵ��������������������������Ȼ�����û���
'�� WritePrivateProfileString ��ͨ������ֵ�Ƿ�Ϊ 0 ���ж��Ƿ�ɹ��ģ�
'���Կ���ͨ���ж� WritePrivateProfileString �ķ���ֵ�Ƿ�� 0 ������ True �� False��

Public Function WriteString(Section As String, key As String, Value As String) As Boolean   'д�ļ�
    WriteString = False
    If NoIniFile() Then
        Exit Function
    End If
    If WritePrivateProfileString(Section, key, Value, IniFileName) = 0 Then
        ErrorMsg = "д��ʧ��"
        Exit Function
    End If
    WriteString = True
End Function

'��������� INI �ļ��ж�ȡһ����ֵ����Ϊ�ַ������ء�
'������� Size �����Ĵ�С�����������ܷ���������ֵ�������������κ���ʾ��
'д��������Ĺؼ��� ReturnStr �ĳ�ʼ����ȡֵ�ϡ�
'VB ���ǲ���Ҫ���ַ������г�ʼ���ģ�Ҳ����Ҫ����ռ䡣
'�������������������ʼ��Ϊһ���㹻�����ַ������Ͳ�����ȷ���ؽ����
'��� C ���Ե��ַ����йأ��Ͳ���˵�ˡ�
'ReturnStr ��ȡֵҲ��Ҫ��Ȥ��Ҫʹ�� Left() ��������ضϡ�
'������ضϣ�ȡ�õĽ���ַ����ͻ��� Size ��ô����
'����ȡ�õ�ֵ���⣬���ಿ�ֶ����ÿո����ġ�
'��ԭ����ǰ��һ����ͬ���� C ���Ե��ַ����йء�
'��Ȼ Left() ����Ҳ����ʹ�� Trim() ���棬Ч����һ���ġ�
Public Function ReadString(Section As String, key As String, Size As Integer) As String
   On Error GoTo Errmsg
    Dim ReturnStr As String
    Dim ReturnLng As Long
    ReadString = vbNullString
    If NoIniFile() Then
        Exit Function
    End If
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(Section, key, vbNullString, ReturnStr, Size, IniFileName)
    ReadString = Left(ReturnStr, ReturnLng)
    Exit Function
Errmsg:
    MsgBox Err.Description
    
End Function


'��������� INI �ļ��ж�ȡһ������ֵ��ʧ��ʱ���� 0��
'���ǵ�ĳЩ����ֵҲ����Ϊ 0����Ӧ��� ErrorMsg �ж��Ƿ�ɹ���
'��������е��������� GetPrivateProfileInt��ΪʲôҪ�����أ�
'��Ϊ GetPrivateProfileInt ����ɹ��򷵻�ȡ�õ�ֵ��������ɹ��򷵻ظ�����Ĭ��ֵ��
'�����ͻ����һ�����������Ҹ���Ĭ��ֵ�� 0��GetPrivateProfileInt ����ȡ�õ�ֵҲ�� 0��
'��ô���ǳɹ�����ʧ���أ�
'ͬ��������Ҹ���Ĭ��ֵ�� 1��GetPrivateProfileInt ����ȡ�õ�ֵҲ�� 1��
'�Ǿ��ǳɹ�����ʧ���أ���Ȼһ��ȡֵ�޷��жϣ��ǾͶ�ȡһ�Σ���һ���趨Ĭ��ֵΪ 0��
'�ڶ����趨Ĭ��ֵΪ 1��INI �ļ�����ֵ��������ҵ�Ĭ��ֵ��ɣ���
'��Ȼ�����鷳һЩ�����Ͼ����������ˡ�
 Public Function ReadInt(Section As String, key As String) As Long
    Dim ReturnLng As Long
    ReadInt = 0
    ReturnLng = GetPrivateProfileInt(Section, key, 0, IniFileName)
    If ReturnLng = 0 Then
        ReturnLng = GetPrivateProfileInt(Section, key, 1, IniFileName)
        If ReturnLng = 1 Then
            ErrorMsg = "���ܶ�ȡ"
            Exit Function
        End If
    End If
    ReadInt = ReturnLng
End Function


