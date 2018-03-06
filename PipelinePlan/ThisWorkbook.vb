
Imports Microsoft.Office.Interop.Excel

Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        Call ProtectSheets()
    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

    End Sub



    Private Sub ThisWorkbook_SheetActivate(Sh As Object) Handles Me.SheetActivate
        Call ProtectSheets()

    End Sub
    Sub ProtectSheets()
        If ActiveSheet.name = "参数" Or ActiveSheet.name = "运行方案" Or ActiveSheet.name = "批次" Or ActiveSheet.name = "下载方案" Then
            Protect([structure]:=True)
        End If
    End Sub
End Class
