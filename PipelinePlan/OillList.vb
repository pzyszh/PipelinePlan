Public Class OilList
    Structure Oil
        Dim Name As String '品名
        Dim Code As Integer '代号
        Dim Color As Integer '颜色
        Dim Density As Decimal '密度
    End Structure
    Private arrOil() As Oil

    Public Sub New(ByVal Oilarr As Application.range)
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

    End Sub
End Class


