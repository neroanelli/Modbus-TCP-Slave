Attribute VB_Name = "Method"
Option Base 1
Public Const LB_FINDSTRING = &H18F
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long  '�ж�����Ϊ��
Public Declare Function timeGetTime Lib "winmm.dll" () As Long             '��ȡ���������ȥ����ʱ��
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long     'ʱ��ֱ���
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
'дINI�ļ�
Public Function WriteFile_INI(Section As String, key As String, Value As String) As Boolean
Dim OpcFile As CIniFile
On Error GoTo ErrHandle
Set OpcFile = New CIniFile
OpcFile.SpecifyIni (App.Path + "\ModbusCfg.ini")
Dim msso As Boolean
msso = OpcFile.WriteString(Section, key, Value)
WriteFile_INI = msso
Exit Function

ErrHandle:
MsgBox err.Description + "OPEN File"
End Function


'��ȡINI�ļ�
Public Function ReadFile_INI(Section As String, key As String) As String
Dim OpcFile As CIniFile
On Error GoTo ErrHandle
Set OpcFile = New CIniFile
OpcFile.SpecifyIni (App.Path + "\ModbusCfg.ini")
Dim msso As String
msso = OpcFile.ReadString(Section, key, 80)
ReadFile_INI = msso
Exit Function

ErrHandle:
MsgBox err.Description + "OPEN File"
End Function

'Log��¼

Public Sub WriteLog(ErrStr As String)

Open App.Path + "\Log.txt" For Append As #1
 'Print #1, vbCrLf$
 Print #1, Now & ":" & ErrStr
Close #1

End Sub

