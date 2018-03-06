VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CINTFCList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'��ͷ����
'Public valOil1 As Integer 'ǰ��
'Dim valOil2 As Integer '����
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
Private strIntfc As String '��ͷ�ַ���
'Private arrIntfc() As typIntfc  '��ͷ����
Private arrIntfc() As New CInterface   '��ͷ����
Private arrUpPlan() As typPlan '�ƻ�������(��������,��������
Private valUpCode As Integer '�½׶��������
Private arrNxtFlow() As Single  '���׶�����
Private valUploaded As Long '��ɵ�����
Private clsPipeProp As CPipeProp '�ܵ�����
Private valActRR As Integer '��ǰ��
Private valActCC As Integer '��ͷ��
Private Const DEAL_PASSROW = 30  '��ͷ����֮��վ��
Private Const DEAL_AHEAD = 31 '��ͷ����֮ǰ��
Private Const DEAL_BEHINE = 32 '��ͷ����֮����
Private Const DEAL_NXT = 33 '��ͷ����֮��վ
Private Const DEAL_DISTANCE = 34 '��ͷ����֮����վ
Private Const DEAL_TMP = 35




Private Sub Class_Initialize()
    Dim tmpOIF As New CInterface
    valActRR = ActiveCell.Row
    Set clsPipeProp = New CPipeProp
    'ȡ��ͷ�б��ַ���
    'Set clsOIQueue = New COI_Queue
    
    valActCC = clsPipeProp.stNum * 2 + 6
    '��ʼ������Դ�
    If valActRR = RUN_DATA_ROW Then
'        strIntfc = Sheets("����").Cells(OIL_INTFC_ROW - 1, OIL_INTFC_COL + 4) '��һ��ȡ�������е���ͷ�б�
    With Sheets("����")
'����ʼ��Ʒ
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
        
    'ȡע���б�
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
MsgBox "�趨Ĭ����Ʒ"
    'ȡ��ע���б�
    For i = RUN_DATA_ROW To valActRR - 1
        valUploaded = valUploaded + ActiveSheet.Cells(i, RUN_DATA_COL).Value
    Next i
    
    
End Sub

Private Function readIntfc() '��ȡ��ͷ�ַ���
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
        For i = a - 1 To 0 Step -1 '�Ӻ���ǰװ����ͷ����
            k = Int(i / 4)
'            if arrIntfc(k)=nothing Set arrIntfc(k) = New CInterface
            Select Case i Mod 4
                '                Case 5 '��ʱ����
                '                    arrIntfc(k).ID = tmp(i)
                Case 0
                    arrIntfc(k).AheadOilColor = tmp(i)
                Case 1 '��Ʒ
                    arrIntfc(k).BehindOilColor = tmp(i)
                Case 2
                    arrIntfc(k).NxtStation = tmp(i)
                Case 3 'վ��\����
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

Private Sub readFlow() '������
    cc = clsPipeProp.stNum
    ReDim arrNxtFlow(cc)
    For i = 0 To cc - 1
        arrNxtFlow(i) = Sheets("����").Cells(valActRR, cc + 3 + i)
        aChildren = clsPipeProp.aChildren(i + 1)
        For j = 0 To UBound(aChildren) - 1
            If aChildren(j) > -1 And i > 0 Then arrNxtFlow(i) = Sheets("����").Cells(valActRR, cc + 2 + aChildren(j)) + arrNxtFlow(i)
        Next
        
    Next
End Sub


Public Sub GoNext()  'ȷ�����׶�����
    Dim tmpUpPlan As Long
    Dim NxtUpload, valUpCode '�½׶μƻ�ע����
    Dim arrNxtPass(1) '�½׶�ע��������վ��ͷ���
    Dim ReasonFlag As Byte
    i = 0
    '������ؽ׶�
    Do While tmpUpPlan <= valUploaded
        If i <= UBound(arrUpPlan) Then
            tmpUpPlan = tmpUpPlan + arrUpPlan(i).Volume
            
        Else
            tmpUpPlan = 0
            Exit Do
        End If
        i = i + 1
    Loop
    valUpCode = i - 1 '���׶����
    NxtUpload = tmpUpPlan - valUploaded
    
    '�����Ʒ�ֶ�
    If readIntfc = -1 Then
        Nxt_pass = -1
        tt = MsgBox("����ͷ��")
        Exit Sub
    End If
    
    Call readFlow
    '�����ͷ��վ��������������ȡ��Сֵ
    With Sheets("����") '����
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
        rngIntfc.Sort key1:="��վ", key2:="����վ", Header:=xlYes
        For i = 0 To UBound(arrIntfc) '������ͷ����վʱ����ע����
            If arrNxtFlow(.Cells(i + 2, DEAL_NXT)) = 0 Then
                .Cells(2 + i, DEAL_TMP) = 999999
            Else
                .Cells(2 + i, DEAL_TMP) = .Cells(2 + i, DEAL_DISTANCE) / (arrNxtFlow(.Cells(i + 2, DEAL_NXT))) * arrNxtFlow(0)
            End If
        Next
        tmpUpPlan = Application.WorksheetFunction.Min(rngIntfc.Columns(6))
        
        If tmpUpPlan < NxtUpload Then '����ͷ��վ�ֶ�
            NxtUpload = tmpUpPlan
            ReasonFlag = 0
        Else
            '�����ر仯����վ
            ReasonFlag = 1
            
            
        End If
        '********************************
        '***********'���¾���վ***********
        For i = .Cells(1, DEAL_PASSROW).CurrentRegion.Rows.Count To 2 Step -1
            If .Cells(i, DEAL_DISTANCE) >= 10 Then
                .Cells(i, DEAL_PASSROW) = valActRR
                .Cells(i, DEAL_DISTANCE) = .Cells(i, DEAL_DISTANCE) - NxtUpload * (arrNxtFlow(.Cells(i, DEAL_NXT))) / arrNxtFlow(0)
            Else
                .Cells(i, DEAL_DISTANCE) = 0
            End If
            If .Cells(i, DEAL_DISTANCE) = 0 Then  '����ͷ��վʱ
                Children = clsPipeProp.Children(.Cells(i, DEAL_NXT))  '��վ
                k = UBound(Children)
                If k = 0 And Children(0) = "-1" Then 'Ϊĩվ��ɾ��
                    .Range(.Cells(i, DEAL_PASSROW), .Cells(i, DEAL_TMP)).Delete SHIFT:=xlUp
                Else  '����ĩվ
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
        '*****������������ͷ����        ******
        Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        rngIntfc.Sort key1:="��վ", key2:="����վ", key3:="��վ��", ORDER3:=xlDescending, Header:=xlYes
        rr = .Cells(1, DEAL_PASSROW).End(xlDown).Row
        For i = rr To 1
            
            If .Cells(i - 1, DEAL_NXT) = .Cells(i, DEAL_NXT) And Abs(.Cells(i - 1, DEAL_DISTANCE) - Cells(i, DEAL_DISTANCE)) < 10 Then
                '����ͬһλ�õ���ͷ
                .Range(.Cells(i, DEAL_PASSROW), .Cells(i, DEAL_TMP)).Delete SHIFT:=xlUp
            End If
        Next
        
        
        rr = .Cells(1, DEAL_PASSROW).End(xlDown).Row
        If ReasonFlag = 1 And valcode < UBound(arrUpPlan) Then
            '��ע��仯��������ͷ
            rr = rr + 1
            .Cells(rr, DEAL_PASSROW) = valActRR
            .Cells(rr, DEAL_AHEAD) = arrUpPlan(valUpCode).OilCode
            .Cells(rr, DEAL_BEHINE) = arrUpPlan(valUpCode + 1).OilCode
            .Cells(rr, DEAL_NXT) = 2 'ĩվ��Ϊ2
            .Cells(rr, DEAL_DISTANCE) = clsPipeProp.Distance(.Cells(rr, DEAL_NXT))
            Sheets("����").Cells(valActRR, 2).Interior.Color = clsPipeProp.oilColor(arrUpPlan(valUpCode).OilCode)
        End If
        ' Set rngIntfc = .Cells(1, DEAL_PASSROW).CurrentRegion
        strIntfc = ""
        For i = 2 To rr
            For j = DEAL_PASSROW To DEAL_DISTANCE
                strIntfc = strIntfc & CInt(.Cells(i, j).Value) & ","
            Next
            
            
        Next
    End With
    Sheets("����").Cells(valActRR, valActCC) = strIntfc
    Sheets("����").Cells(valActRR, 2) = NxtUpload
End Sub

Private Sub Class_Terminate()
    Set clsPipeProp = Nothing
End Sub


