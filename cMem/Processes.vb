Imports System.Runtime.InteropServices
Public Class ProcessList
    Public Processes As New Dictionary(Of Integer, Process)
    Sub New(Optional ByVal process_name As String = "", Optional ByVal directx_only_processes As Boolean = False, Optional ByVal filters As String = "")
        Dim non_filtered As Process() = Process.GetProcesses
        Try
            For Each p As Process In non_filtered
                If p.Id <> 4 Then 'Process ID 4 is for system and is not actually a process
                    If directx_only_processes Then
                        Try
                            If p.ProcessName.ToLower.Contains(process_name.ToLower) Then
                                If isValid(p) Then
                                    For Each modu As ProcessModule In p.Modules
                                        If LCase(Left(modu.ModuleName, 4)) = "d3dx" Then
                                            Processes.Add(p.Id, p)
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Catch ex As Exception
                            ' Console.WriteLine("Process: " & p.ProcessName & "Error: " & Marshal.GetLastWin32Error())
                        End Try
                    Else
                        If p.ProcessName.ToLower.Contains(process_name.ToLower) Then
                            Processes.Add(p.Id, p)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("Process Win32 Error: " & Marshal.GetLastWin32Error())
        End Try

    End Sub
    Private Function isValid(ByVal p As Process) As Boolean
        Try
            Dim hToken As IntPtr
            If NativeApi.OpenProcessToken(p.Handle, Security.Principal.TokenAccessLevels.Query, hToken) Then
                Using wi As New Security.Principal.WindowsIdentity(hToken)
                    'Console.WriteLine("process: " & p.ProcessName & " owner:" & wi.Name)
                    If p.ProcessName = "devenv" Then Return False
                    If wi.IsSystem Then
                        Return False
                    Else
                        Return True
                    End If

                End Using
            End If
            Return True
        Catch ex As Exception
            Return True
        End Try
    End Function
End Class
