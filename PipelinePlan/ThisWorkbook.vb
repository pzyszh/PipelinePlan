
Imports Microsoft.Office.Interop.Excel

Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

    End Sub



    Private Sub ThisWorkbook_SheetActivate(Sh As Object) Handles Me.SheetActivate
        If Sh.name="参数" Then
            Me.Protect()
        End If
    End Sub
End Class
