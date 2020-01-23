Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Diagnostics
Imports System.Threading

Public Class ThreadData
    Private mem As Memory
    Sub New(ByRef memref As Memory)
        mem = memref
    End Sub
    Public Function GetPlatform() As NativeApi.Platform
        Dim sysInfo As New NativeApi.SYSTEM_INFO()

        If System.Environment.OSVersion.Version.Major > 5 OrElse (System.Environment.OSVersion.Version.Major = 5 AndAlso System.Environment.OSVersion.Version.Minor >= 1) Then
            NativeApi.GetNativeSystemInfo(sysInfo)
        Else
            NativeApi.GetSystemInfo(sysInfo)
        End If

        Select Case sysInfo.wProcessorArchitecture
            Case NativeApi.PROCESSOR_ARCHITECTURE_IA64, NativeApi.PROCESSOR_ARCHITECTURE_AMD64
                Return NativeApi.Platform.X64
            Case NativeApi.PROCESSOR_ARCHITECTURE_INTEL
                Return NativeApi.Platform.X86
            Case Else
                Return NativeApi.Platform.Unknown
        End Select
    End Function
    Public ThreadList As New Dictionary(Of Integer, ThreadInfo)
    Public Function SuspendThread(ByVal thread_id As Long) As Boolean
        Dim PtrThread As Long = NativeApi.OpenThread(&H1F03FF, False, thread_id) 'Open the thread for terminating
        NativeApi.SuspendThread(PtrThread)
        NativeApi.CloseHandle(PtrThread) 'close the handle we opened on the thread
    End Function
    Public Function TerminateThread(ByVal thread_id As Long) As Boolean
        Dim PtrThread As Long = NativeApi.OpenThread(&H1F03FF, False, thread_id) 'Open the thread for terminating
        NativeApi.TerminateThread(PtrThread, 0)
        NativeApi.CloseHandle(PtrThread) 'close the handle we opened on the thread
    End Function
    Public Function ResumeThread(ByVal thread_id As Long) As Boolean
        Dim PtrThread As Long = NativeApi.OpenThread(&H1F03FF, False, thread_id) 'Open the thread for terminating
        NativeApi.ResumeThread(PtrThread)
        NativeApi.CloseHandle(PtrThread) 'close the handle we opened on the thread
    End Function
    Public Function GetThreads() As Boolean
        Stop
        Dim processThreads As ProcessThreadCollection = Nothing
        ThreadList.Clear()
        For Each pt As ProcessThread In mem.Process_Obj.Threads
            Dim PtrThread As Long = NativeApi.OpenThread(&H1F03FF, False, pt.Id) 'Open the thread for terminating
            Dim status As Long 'Windows NT Status Code
            Dim info As New NativeApi.THREAD_BASIC_INFORMATION 'Thread Basic Information -- used if type 0 used in NtQueryInformationThread
            Dim start_addr As Long = 0 'Defines the start address as 0
            Dim m As ProcessModule
            status = NativeApi.NtQueryInformationThread(PtrThread, 9, start_addr, 4, IntPtr.Zero)
            m = findModule(start_addr)
            Dim td As New ThreadInfo
            td.id = pt.Id
            td.start_address = start_addr
            td.status = status
            td.state = pt.ThreadState
            If m IsNot Nothing Then
                td.display = m.ModuleName & "+" & Hex(start_addr - Long.Parse(m.BaseAddress))
                td.module_addr = m.BaseAddress
                td.module_name = m.ModuleName
            Else
                td.display = start_addr.ToString
            End If
            NativeApi.CloseHandle(PtrThread) 'close the handle we opened on the thread
            ThreadList.Add(pt.Id, td)
        Next
        Return Nothing
    End Function
    Public Function findModule(ByVal addr As Long) As ProcessModule
        Try
            For Each modu As ProcessModule In mem.Process_Obj.Modules
                If addr > Long.Parse(modu.BaseAddress) And addr < Long.Parse(modu.ModuleMemorySize) + Long.Parse(modu.BaseAddress) Then
                    Return modu
                End If
            Next
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
Public Class ThreadInfo
    Public id As Long = 0
    Public start_address As Long = 0
    Public status As Long = 0
    Public state As Diagnostics.ThreadState
    Public module_name As String = ""
    Public module_addr As Long = 0
    Public display As String = ""
End Class
