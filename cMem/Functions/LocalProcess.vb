Public Class LocalProcess
    Public Shared Function ModuleObj(ByVal name As String) As ProcessModule
        For Each modu As ProcessModule In Process.GetCurrentProcess.Modules
            If LCase(Left(modu.ModuleName, name.Length)) = LCase(name) Then
                Return modu
            End If
        Next
        Return Nothing
    End Function
End Class
