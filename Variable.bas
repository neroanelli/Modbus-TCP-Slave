Attribute VB_Name = "Variable"
Type TagData
    Tag As String       '结构体中标签名
    TagName As String
    HH As Double            '结构体中量程上限
    LL As Double            '结构体中量程下限
    N As Integer            '结构体中有效采集累加的计数值
    Value As Double         '对应的采集值
    
End Type


Public Data() As TagData
Public TagSum As Integer
