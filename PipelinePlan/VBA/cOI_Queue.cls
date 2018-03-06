VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COI_Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim arr(100) As New CInterface
Dim front As Integer
Dim rear  As Integer

Sub outQ()
    If rear = fornt Then
        Call QueueEmpty
    Else
        If rear = 0 Then rear = 99 Else rear = rear - 1
    End If
End Sub

Sub inQ(ByRef x As CInterface)
    If (rear + 1) Mod 100 = fornt Then
        Call QueueFull
    Else
        If rear = 99 Then rear = 0 Else rear = rear + 1
    End If
    Set arr(rear) = x
    If (rear - front + 100) Mod 100 Then Call CheckQueue '����������ͷ����һ��ʱ�������غ���ͷ
End Sub

Sub MoveQ(ByVal V As Long)  '��ͷǰ��V����
    For i = front To rear
        arr(i).Distance = arr(i).Distance - V
        If -10 < arr(i).Distance < 10 Then
            arr(i).Distance = 0
        ElseIf arr(i).Distance < 0 Then
            Call myError(3)
        End If
    Next
End Sub
Sub InitQ()
    rear = front
End Sub
Private Sub Class_Initialize()
    front = 0
    rear = 0
End Sub

Private Sub QueueFull()
    MsgBox ("��������")
    End
End Sub

Private Sub QueueEmpty()
    MsgBox ("Ϊ��")
    End
End Sub

Private Sub myError(ByVal x As Integer)
    Select Case x
        Case 0
            MsgBox ("δ֪����")
            End
        Case 1
            MsgBox ("��ͷ��������")
            End
        Case 2
            MsgBox ("��ͷ����Ϊ��")
            End
        Case 3
            MsgBox ("��ͷ��Խվ")
            End
    End Select
End Sub
Private Sub CheckQueue()
    'MsgBox "CleanQueue��ûд"
    
    rl = (rear + 99) Mod 100
    If arr(rear).Distance = 0 Then
        arr(rl).BehindOilColor = arr(rear).BehindOilColor
        rear = rl
    End If
    
End Sub