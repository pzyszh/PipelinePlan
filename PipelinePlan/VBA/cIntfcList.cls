VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CINTFCList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'油头处理
'Public valOil1 As Integer '前行
'Dim valOil2 As Integer '后行
'Dim valNxtSta As Byte
'Dim valDistance As Long
Private Type typPlan
    OilCode As Long
    Volume As Single
End Type
'Private Type typIntfc
'    AheadOilColor As Long
'    BehindOilColor As Long
'    NxtStation As Long
'    Distance As Single
'    ID As Single
'End Type
Private strIntfc As String '油头字符串
'Private arrIntfc() As typIntfc  '油头数组
Private arrIntfc() As New CInterface   '油头数组
Private arrUpPlan() As typPlan '计划的上载(上载种类,上载量）
Private valUpCode As Integer '下阶段上载序号
Private arrNxtFlow() As Single  '本阶段流量
Private valUploaded As Long '完成的上载
Private clsPipeProp As CPipeProp '管道参数
Private valActRR As Integer '当前行
Private valActCC As Integer '油头列
Private Const DEAL_PASSROW = 30  '油头处理之过站行
Private Const DEAL_AHEAD = 31 '油头处理之前行
Private Const DEAL_BEHINE = 32 '油头处理之后行
Private Const DEAL_NXT = 33 '油头处理之下站
Private Const DEAL_DISTANCE = 34 '油头处理之距下站
Private Const DEAL_TMP = 35




Private Sub Class_Initialize()
    Dim tmpOIF As New CInterface
    valActRR = ActiveCell.Row
    Set clsPipeProp = New CPipeProp
    '取油头列表字符串
    'Set clsOIQueue = New COI_Queue
    
    valActCC = clsPipeProp.stNum * 2 + 6
    '初始行特殊对待
    If valActRR = RUN_DATA_ROW Then
'        strIntfc = Sheets("参数").Cells(OIL_INTFC_ROW - 1, OIL_INTFC_COL + 4) '第一行取参数表中的油头列表
    With Sheets("参数")
'读初始油品
        mRow = OIL_INTFC_ROW
        mCol = OIL_INTFC_COL
        re = .Cells(mRow, mCol).End(xlDown).Row - mRow
        For i = 0 To re
            tmpOIF.NxtStation = clsPipeProp.StCode(.Cells(mRow + i, mCol + 2))
            tmpOIF.BehindOilColor = clsPipeProp.OilCode(.Cells(mRow + i, mCol))
            vv = .Cells(mRow + i, mCol + 1).Value
            If vv > 0 Then tmpOIF.Distance = vv Else tmpOIF.Distance = clsPipeProp.Distance(tmpOIF.NxtStation)
            clsOIQueue(tmpOIF.NxtStation).inQ tmpOIF
        Next
        
    '取注入列表
        mRow = OIL_UPLOAD_ROW
        mCol = OIL_UPLOAD_COL
        re = .Cells(mRow, mCol).End(xlDown).Row - mRow
        ReDim arrUpPlan(0 To re)
        For i = 0 To re
            tmpOIF.NxtStation = 1
            tmpOIF.BehindOilColor = clsPipeProp.OilCode(.Cells(mRow + i, mCol))
            tmpOIF.Distance = .Cells(mRow + i, mCol + 2).Value
            clsOIQueue(tmpOIF.NxtStation).inQ tmpOIF
        Next
    End With
        
        
    Else
        
        cc = valActCC
        strIntfc = ActiveSheet.Cells(valActRR - 1, cc)

    End If
MsgBox "设定默认油品"
    '取已注入列表
    For i = RUN_DATA_ROW To valActRR - 1
        valUploaded = valUploaded + ActiveSheet.Cells(i, RUN_DATA_COL).Value
    Next i
    
    
End Sub

Private Function readIntfc() '读取油头字符串
    Dim tmp
    Dim a
    
    tmp = Split(strIntfc, ",")
    a = UBound(tmp)
    If a = -1 Then
        Erase arrIntfc
        readIntfc = -1
    Else
        n = a / 4 - 1
        ReDim arrIntfc(Int(n))
        For i = a - 1 To 0 Step -1 '从后往前装入油头序列
            k = Int(i / 4)
'            if arrIntfc(k)=nothing Set arrIntfc(k) = New CInterface
            Select Case i Mod 4
                '                Case 5 '暂时无用
                '                    arrIntfc(k).ID = tmp(i)
                Case 0
                    arrIntfc(k).AheadOilColor = tmp(i)
                Case 1 '油品
                    arrIntfc(k).BehindOilColor = tmp(i)
                Case 2
                    arrIntfc(k).NxtStation = tmp(i)
                Case 3 '站场\距离
                    arrIntfc(k).Distance = tmp(i)
            End Select
            'clsPipeProp.clsOIQueue(arrIntfc(k).NxtStation).inQ (arrIntfc(k))
            
            readIntfc = 1
            'Debug.Print k
            'Debug.Print i
        Next
        For i = 0 To n
            clsOIQueue(arrIntfc(i).NxtStation).inQ arrIntfc(i)
        Next
    End If
End Function

Private Sub readFlow() '读流量
    cc = clsPipeProp.stNum
    ReDim arrNxtFlow(cc)
    For i = 0 To cc - 1
        arrNxtFlow(i) = Sheets("运行").Cells(valActRR, cc + 3 + i)
        aChildren = clsPipeProp.aChildren(i + 1)
        For j = 0 To UBound(aChildren) - 1
            If aChildren(j) > -1 And i > 0 Then arrNxtFlow(i) = Sheets("运行").Cells(valActRR, cc + 2 + aChildren(j)) + arrNxtFlow(i)
        Next
        
    Next
End Sub


Public Sub GoNext()  '确定本阶段上载
    Dim tmpUpPlan As Long
    Dim NxtUpload, valUpCode '下阶段计划注入量
    Dim arrNxtPass(1) '下阶段注入量、到站油头序号
    Dim ReasonFlag As Byte
    i = 0
    '检查上载阶段
    Do While tmpUpPlan <= valUploaded
        If i <= UBound(arrUpPlan) Then
            tmpUpPlan = tmpUpPlan + arrUpPlan(i).Volume
            
        Else
            tmpUpPlan = 0
            Exit Do
        End If
        i = i + 1
    Loop
    valUpCode = i - 1 '本阶段序号
    NxtUpload = tmpUpPlan - valUploaded
    
    '检查油品分段
    If readIntfc = -1 Then
        Nxt_pass = -1
        tt = MsgBox("无油头表")
        Exit Sub
    End If
    
    Call readFlow
    '检查油头到站所需上载量，并取最小值
    With Sheets("参数") '排序
        Dim rngIntfc As Range
        Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        rngIntfc.CurrentRegion.Offset(1, 0).Clear
        For i = 0 To UBound(arrIntfc)
            .Cells(i + 2, DEAL_PASSROW) = arrIntfc(i).ID
            .Cells(i + 2, DEAL_AHEAD) = arrIntfc(i).AheadOilColor
            .Cells(i + 2, DEAL_BEHINE) = arrIntfc(i).BehindOilColor
            .Cells(i + 2, DEAL_DISTANCE) = arrIntfc(i).Distance
            .Cells(i + 2, DEAL_NXT) = arrIntfc(i).NxtStation
        Next
        Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        rngIntfc.Sort key1:="下站", key2:="距下站", Header:=xlYes
        For i = 0 To UBound(arrIntfc) '计算油头到下站时所需注入量
            If arrNxtFlow(.Cells(i + 2, DEAL_NXT)) = 0 Then
                .Cells(2 + i, DEAL_TMP) = 999999
            Else
                .Cells(2 + i, DEAL_TMP) = .Cells(2 + i, DEAL_DISTANCE) / (arrNxtFlow(.Cells(i + 2, DEAL_NXT))) * arrNxtFlow(0)
            End If
        Next
        tmpUpPlan = Application.WorksheetFunction.Min(rngIntfc.Columns(6))
        
        If tmpUpPlan < NxtUpload Then '因油头过站分段
            NxtUpload = tmpUpPlan
            ReasonFlag = 0
        Else
            '因上载变化而过站
            ReasonFlag = 1
            
            
        End If
        '********************************
        '***********'更新距下站***********
        For i = .Cells(1, DEAL_PASSROW).CurrentRegion.Rows.Count To 2 Step -1
            If .Cells(i, DEAL_DISTANCE) >= 10 Then
                .Cells(i, DEAL_PASSROW) = valActRR
                .Cells(i, DEAL_DISTANCE) = .Cells(i, DEAL_DISTANCE) - NxtUpload * (arrNxtFlow(.Cells(i, DEAL_NXT))) / arrNxtFlow(0)
            Else
                .Cells(i, DEAL_DISTANCE) = 0
            End If
            If .Cells(i, DEAL_DISTANCE) = 0 Then  '当油头过站时
                Children = clsPipeProp.Children(.Cells(i, DEAL_NXT))  '下站
                k = UBound(Children)
                If k = 0 And Children(0) = "-1" Then '为末站则删除
                    .Range(.Cells(i, DEAL_PASSROW), .Cells(i, DEAL_TMP)).Delete SHIFT:=xlUp
                Else  '不是末站
                    .Cells(i, DEAL_DISTANCE) = clsPipeProp.Distance(Children(0))
                    .Cells(i, DEAL_NXT) = Children(0)
                    For j = 1 To k Step 1
                        rr = .Cells(1, DEAL_PASSROW).End(xlDown).Row
                        .Range(.Cells(i, DEAL_PASSROW), .Cells(i, DEAL_DISTANCE)).Copy (.Cells(rr + 1, DEAL_PASSROW))
                        .Cells(rr, DEAL_DISTANCE) = clsPipeProp.Distance(Children(j))
                        .Cells(rr, DEAL_NXT) = Children(j)
                    Next
                End If
            End If
            
        Next
        '*****整理并生成油头序列        ******
        Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        rngIntfc.Sort key1:="下站", key2:="距下站", key3:="过站行", ORDER3:=xlDescending, Header:=xlYes
        rr = .Cells(1, DEAL_PASSROW).End(xlDown).Row
        For i = rr To 1
            
            If .Cells(i - 1, DEAL_NXT) = .Cells(i, DEAL_NXT) And Abs(.Cells(i - 1, DEAL_DISTANCE) - Cells(i, DEAL_DISTANCE)) < 10 Then
                '清理同一位置的油头
                .Range(.Cells(i, DEAL_PASSROW), .Cells(i, DEAL_TMP)).Delete SHIFT:=xlUp
            End If
        Next
        
        
        rr = .Cells(1, DEAL_PASSROW).End(xlDown).Row
        If ReasonFlag = 1 And valcode < UBound(arrUpPlan) Then
            '因注入变化则增加油头
            rr = rr + 1
            .Cells(rr, DEAL_PASSROW) = valActRR
            .Cells(rr, DEAL_AHEAD) = arrUpPlan(valUpCode).OilCode
            .Cells(rr, DEAL_BEHINE) = arrUpPlan(valUpCode + 1).OilCode
            .Cells(rr, DEAL_NXT) = 2 '末站号为2
            .Cells(rr, DEAL_DISTANCE) = clsPipeProp.Distance(.Cells(rr, DEAL_NXT))
            Sheets("运行").Cells(valActRR, 2).Interior.Color = clsPipeProp.oilColor(arrUpPlan(valUpCode).OilCode)
        End If
        ' Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        strIntfc = ""
        For i = 2 To rr
            For j = DEAL_PASSROW To DEAL_DISTANCE
                strIntfc = strIntfc & CInt(.Cells(i, j).Value) & ","
            Next
            
            
        Next
    End With
    Sheets("运行").Cells(valActRR, valActCC) = strIntfc
    Sheets("运行").Cells(valActRR, 2) = NxtUpload
End Sub

Private Sub Class_Terminate()
    Set clsPipeProp = Nothing
End Sub



