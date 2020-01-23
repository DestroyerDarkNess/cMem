Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Public Class MemRegion
    Public Address As Long
    Public Length As Long
End Class
Public Class Mbi
    Private mem As Memory

    Sub New(ByRef memref As Memory)
        mem = memref
    End Sub
    Public Structure MEMORY_BASIC_INFORMATION
        Public BaseAddress As Integer
        Public AllocationBase As Integer
        Public AllocationProtect As AllocationProtect
        Public RegionSize As Integer
        Public State As Integer
        Public Protect As Integer
        Public Type As Integer
    End Structure
    Public Enum AllocationProtect As Integer
        PAGE_EXECUTE = &H10
        PAGE_EXECUTE_READ = &H20
        PAGE_EXECUTE_READWRITE = &H40
        PAGE_EXECUTE_WRITECOPY = &H80
        PAGE_NOACCESS = &H1
        PAGE_READONLY = &H2
        PAGE_READWRITE = &H4
        PAGE_WRITECOPY = &H8
        PAGE_GUARD = &H100
        PAGE_NOCACHE = &H200
        PAGE_WRITECOMBINE = &H400
    End Enum
    Public Function getWriteableRegions() As Dictionary(Of Integer, MemRegion)
        Dim address As IntXLib.IntX = 0
        Dim dict As New Dictionary(Of Integer, MemRegion)
        Dim m As MEMORY_BASIC_INFORMATION
        Dim i As Integer = 0
        Dim base As IntXLib.IntX = 0
        Dim size As IntXLib.IntX = 0
        Try
            While NativeApi.VirtualQueryEx(mem.Process_Handle, address, m, Marshal.SizeOf(GetType(MEMORY_BASIC_INFORMATION))) <> 0
                If m.AllocationProtect = AllocationProtect.PAGE_EXECUTE_READWRITE Or m.AllocationProtect = AllocationProtect.PAGE_EXECUTE_WRITECOPY Or m.AllocationProtect = AllocationProtect.PAGE_READWRITE Or m.AllocationProtect = AllocationProtect.PAGE_WRITECOMBINE Or m.AllocationProtect = AllocationProtect.PAGE_WRITECOPY Then
                    i += 1
                    '  Console.WriteLine(Hex(m.BaseAddress).ToString & " " & m.RegionSize & " " & m.AllocationProtect.ToString)
                    Dim reg As New MemRegion
                    reg.Address = m.BaseAddress
                    reg.Length = m.RegionSize
                    dict.Add(i, reg)
                End If
                base = m.BaseAddress
                size = m.RegionSize
                address = base + size

                'Console.WriteLine(Hex(address))
            End While
            Return dict
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
            Return Nothing
        End Try
    End Function

End Class
