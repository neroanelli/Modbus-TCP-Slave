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
'Alias "GetPrivateProfileIntA" ( _       ' 返回所读取的长整型值
'    ByVal lpApplicationName As String, _    ' 要读取的段 (Section) 名称
'    ByVal lpKeyName As String, _            ' 要读取的的键 (Key) 名称
'    ByVal nDefault As Long, _               ' 指定默认值，如果读取时出错，则返回该值
'    ByVal lpFileName As String) As Long     ' 指定要读的 INI 文件名
'
'Private Declare Function GetPrivateProfileString Lib "kernel32" _
'Alias "GetPrivateProfileStringA" ( _    ' 返回所读取的字符串值的真实长度
'    ByVal lpApplicationName As String, _    ' 要读取的段 (Section) 名称
'    ByVal lpKeyName As Any, _               ' 要读取的的键 (Key) 名称
'    ByVal lpDefault As String, _            ' 指定默认值，如果读取时出错，则返回该值
'    ByVal lpReturnedString As String, _     ' 指定接收返回值的字符串变量
'    ByVal nSize As Long, _                  ' 指定允许字符串值的最大长度
'    ByVal lpFileName As String) As Long     ' 指定要读的 INI 文件名
'
'Private Declare Function WritePrivateProfileString Lib "kernel32" _
'Alias "WritePrivateProfileStringA" ( _  ' 如果成功返回非 0 值，失败返回 0
'    ByVal lpApplicationName As String, _    ' 要写入的段 (Section) 名称
'    ByVal lpKeyName As Any, _               ' 要写入的的键 (Key) 名称
'    ByVal lpString As Any, _                ' 要写入的值 (Value)，以字符串表示
'    ByVal lpFileName As String) As Long     ' 指定要写的 INI 文件名
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private IniFileName As String    '文件名
Public ErrorMsg As String        '错误信息




Private Sub Class_Initialize()
    IniFileName = vbNullString
    ErrorMsg = vbNullString

End Sub

Public Sub SpecifyIni(FilePathName As String)      '指定 INI 文件名给 CIniFile
    IniFileName = Trim(FilePathName)
End Sub

 Private Function NoIniFile() As Boolean  '判断是否已经指定了 INI 文件名
    NoIniFile = True
    If IniFileName = vbNullString Then
        ErrorMsg = "没有指定 INI 文件"
        Exit Function
    End If
    ErrorMsg = vbNullString
    NoIniFile = False
End Function

'该方法在 INI 文件中写入一个键值，成功返回 True，失败返回 False。
'根据 WritePrivateProfileString 的需要，除了文件名这一参数不用提供之外，
'需要提供段名、键名和值三个参数，而且这三个参数当然来自用户。
'而 WritePrivateProfileString 是通过返回值是否为 0 来判断是否成功的，
'所以可以通过判断 WritePrivateProfileString 的返回值是否非 0 来返回 True 或 False。

Public Function WriteString(Section As String, key As String, Value As String) As Boolean   '写文件
    WriteString = False
    If NoIniFile() Then
        Exit Function
    End If
    If WritePrivateProfileString(Section, key, Value, IniFileName) = 0 Then
        ErrorMsg = "写入失败"
        Exit Function
    End If
    WriteString = True
End Function

'这个方法在 INI 文件中读取一个键值，作为字符串返回。
'如果参数 Size 给定的大小不够，将不能返回完整的值串，但不会有任何提示。
'写这个函数的关键在 ReturnStr 的初始化和取值上。
'VB 中是不需要对字符串进行初始化的，也不需要分配空间。
'但是这里如果不将它初始化为一个足够长的字符串，就不能正确返回结果。
'这和 C 语言的字符串有关，就不多说了。
'ReturnStr 的取值也需要有趣，要使用 Left() 函数将其截断。
'如果不截断，取得的结果字符串就会有 Size 那么长，
'除了取得的值以外，其余部分都是用空格填充的。
'其原因与前面一点相同，与 C 语言的字符串有关。
'当然 Left() 函数也可以使用 Trim() 代替，效果是一样的。
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


'这个方法在 INI 文件中读取一个整数值，失败时返回 0。
'考虑到某些键的值也可能为 0，故应结合 ErrorMsg 判断是否成功。
'这个方法中调用了两次 GetPrivateProfileInt，为什么要这样呢？
'因为 GetPrivateProfileInt 如果成功则返回取得的值，如果不成功则返回给定的默认值。
'这样就会出现一种情况：如果我给的默认值是 0，GetPrivateProfileInt 函数取得的值也是 0，
'那么它是成功还是失败呢？
'同样，如果我给的默认值是 1，GetPrivateProfileInt 函数取得的值也是 1，
'那就是成功还是失败呢？既然一次取值无法判断，那就多取一次，第一次设定默认值为 0，
'第二次设定默认值为 1，INI 文件的中值不会跟着我的默认值变吧？！
'虽然这样麻烦一些，但毕竟把问题解决了。
 Public Function ReadInt(Section As String, key As String) As Long
    Dim ReturnLng As Long
    ReadInt = 0
    ReturnLng = GetPrivateProfileInt(Section, key, 0, IniFileName)
    If ReturnLng = 0 Then
        ReturnLng = GetPrivateProfileInt(Section, key, 1, IniFileName)
        If ReturnLng = 1 Then
            ErrorMsg = "不能读取"
            Exit Function
        End If
    End If
    ReadInt = ReturnLng
End Function



