VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPipeProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type typStation
    Name As String
    Distance As Long
    Children As String
    outside_oil As Integer
'    clsOIQueue As CIntfcQueue
End Type
'Public clsOIQueue(100) As CIntfcQueue
Private Type typOil
    Name As String
    Color As Long
    Density As Single
End Type

Private pTree() As typStation
Dim arrOil() As typOil
Public stNum As Integer '站场数
Public oilNum As Integer '油品种类数


Private Sub Class_Initialize()
    '***********读管道参数***************
    With Sheets("参数")
        Dim st_Row '数据起始行
        st_Row = 3
        re = .Cells(st_Row, 1).End(xlDown).Row
        n = re - st_Row + 1 '总节点数
        ReDim pTree(1 To n)
        ReDim clsOIQueue(1 To n)
        
        For i = 1 To re - st_Row + 1
            pTree(i).Distance = .Cells(i + st_Row - 1, 4)
            pTree(i).Children = .Cells(i + st_Row - 1, 3)
            pTree(i).Name = .Cells(i + st_Row - 1, 2)
            Set clsOIQueue(i) = New COI_Queue '初始化油头队列
        Next i
        
        stNum = n
        '***读油品参数*********************
        Dim mRow, mCol
        mRow = 3
        mCol = 6
        re = .Cells(mRow, mCol).End(xlDown).Row - mRow
        ReDim arrOil(1 To re + 1)
        For i = 1 To re + 1
            arrOil(i).Name = .Cells(mRow + i - 1, mCol + 1)
            arrOil(i).Color = .Cells(mRow + i - 1, mCol + 3)
            arrOil(i).Density = .Cells(mRow + i - 1, mCol + 4)
        Next
        oilNum = re + 1
    End With
End Sub
'    Property Get node(n As Integer) As typStation
'        node = pTree(i)
'    End Property
Private Sub TraDistance(n, tresult)    '树，开始节点号，结果
    Dim aChildren
    'Dim tresult As typStation
    aChildren = Split(pTree(n).Children, ",")
    For i = 0 To UBound(aChildren)
        If aChildren(i) > -1 Then Call TraDistance(CInt(aChildren(i)), tresult)
    Next
    tresult = tresult + pTree(n).Distance
    
End Sub
Private Sub TraChildren(n, tresult)    '树，开始节点号，结果
    Dim aChildren
    aChildren = Split(pTree(n).Children, ",")
    For i = 0 To UBound(aChildren)
        If aChildren(i) > -1 Then Call TraChildren(CInt(aChildren(i)), tresult)
    Next
    If pTree(n).Children <> "-1" Then tresult = tresult & pTree(n).Children & ","
    
End Sub
'Function sumChildren(n) '
'    Dim s
'   ' Call Traversal(N, s)
'    sumChildren = Traversal(n) - pTree(n).Distance
'End Function
Property Get StCode(strName) '取站场代码
    For i = 1 To stNum
        If pTree(i).Name = strName Then
            StCode = i
            Exit For
        End If
    Next
    If StCode = "" Then Call myError(strName)
End Property
Property Get oilName(x) '取油品名称
    Dim i
    If 0 < x And x < oilNum Then
        oilName = arrOil(x).Name
    Else
        For i = 1 To oilNum
            If arrOil(i).Color = x Then oilName = arrOil(i).Name
        Next
    End If
    If oilName = "" Then Call myError(x)
End Property
Property Get oilColor(x) '取油品颜色代码
    Dim i
    If TypeName(x) = "String" Then
        For i = 1 To oilNum
            If arrOil(i).Name = x Then oilColor = arrOil(i).Color
        Next
    ElseIf 0 < x And x < oilNum Then
        oilColor = arrOil(x).Color
    End If
    
    If oilColor = "" Then Call myError(x)
End Property
Property Get OilCode(x) '取油品代码
    Dim i
    For i = 1 To oilNum
        If arrOil(i).Name = x Then
            OilCode = i
            Exit For
        ElseIf arrOil(i).Color = x Then
            OilCode = i
            Exit For
        End If
    Next
    
    If OilCode = "" Then Call myError(x)
End Property
Property Get oilDensity(x) '取油品密度
    Dim i
    If TypeName(x) = "String" Then
        For i = 1 To oilNum
            If arrOil(i).Name = x Then oilDensity = arrOil(i).Density
        Next
    ElseIf 0 < x And x < oilNum Then
        oilDensity = arrOil(x).Density
    Else
        For i = 1 To oilNum
            If arrOil(i).Color = x Then oilDensity = arrOil(i).Density
        Next
    End If
    
    If oilDensity = "" Then Call myError(x)
End Property
Property Get Children(ByVal n As Integer)
    If n < 1 Or n > stNum Then Call myError(n)
    Children = Split(pTree(n).Children, ",")
End Property
Property Get aChildren(ByVal n As Integer)
    Call TraChildren(n, result)
    aChildren = Split(result, ",")
End Property
Property Get Distance(ByVal n As Integer)
    If n < 0 Or n > stNum Then Call myError(n)
    Distance = pTree(n).Distance
End Property
Private Sub myError(x)
    MsgBox ("错误的输入" & x)
    End
End Sub

Private Sub TraTree(aNode, n, tresult) '树，开始节点号，结果
    Dim aChildren
    aChildren = Split(aNode(n, 2), ",")
    For i = 0 To UBound(aChildren)
        Call TraTree(aNode, CInt(aChildren(i)), tresult)
    Next
    tresult = tresult + aNode(n, 1)
End Sub

'Property Get OIQueue(ByVal n As Integer)
'    If n < 1 Or n > stNum Then Call myError(n)
'   Set OIQueue = pTree(n).clsOIQueue
'End Property
