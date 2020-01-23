Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Public Class Memory
    Public Process_Handle As Long
    Public Process_Obj As Process
    Public Process_OpenedHandle As IntPtr
    Public Threads As ThreadData
    Public MemBasicInfo As Mbi
    Sub New()
        AdjustPriv() 'adjust our own priveleges if possible
        Threads = New ThreadData(Me)
        MemBasicInfo = New Mbi(Me)
    End Sub
    Public Sub Attach(ByVal proc As Process)
        Try
            Process_Handle = proc.Handle
            Process_Obj = proc
            Dim access As UInt32 = NativeApi.ProcessAccessType.PROCESS_VM_READ Or NativeApi.ProcessAccessType.PROCESS_VM_WRITE Or NativeApi.ProcessAccessType.PROCESS_VM_OPERATION Or NativeApi.ProcessAccessType.PROCESS_CREATE_THREAD
            Process_OpenedHandle = NativeApi.OpenProcess(access, 1, CUInt(Process_Obj.Id))
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub
#Region "Adjust Privelages"
    Public Sub AdjustPriv()
        Dim lastWin32Error As Integer = 0

        'Get the LUID that corresponds to the Shutdown privilege, if it exists.
        Dim luid_Shutdown As NativeApi.LUID
        If Not NativeApi.LookupPrivilegeValue(Nothing, "SeDebugPrivilege", luid_Shutdown) Then
            lastWin32Error = Marshal.GetLastWin32Error()
            Console.WriteLine("Error looking up privelages: " & lastWin32Error.ToString)
        End If

        'Get the current process's token.
        Dim hProc As IntPtr = Process.GetCurrentProcess().Handle
        Dim hToken As IntPtr
        If Not NativeApi.OpenProcessToken(hProc, NativeApi.Tokens.ADJUST_PRIVILEGES Or NativeApi.Tokens.QUERY, hToken) Then
            lastWin32Error = Marshal.GetLastWin32Error()
            Console.WriteLine("Error Opening Process Tokens: " & lastWin32Error.ToString)
        End If


        'Set up a LUID_AND_ATTRIBUTES structure containing the Shutdown privilege, marked as enabled.
        Dim luaAttr As New NativeApi.LUID_AND_ATTRIBUTES
        luaAttr.Attributes = NativeApi.SE_PRIVILEGE_ENABLED

        'Set up a TOKEN_PRIVILEGES structure containing only the shutdown privilege.
        Dim newState As New NativeApi.TOKEN_PRIVILEGES
        newState.PrivilegeCount = 1
        newState.Privileges = New NativeApi.LUID_AND_ATTRIBUTES() {luaAttr}

        'Set up a TOKEN_PRIVILEGES structure for the returned (modified) privileges.
        Dim prevState As NativeApi.TOKEN_PRIVILEGES = New NativeApi.TOKEN_PRIVILEGES
        ReDim prevState.Privileges(CInt(newState.PrivilegeCount))

        'Apply the TOKEN_PRIVILEGES structure to the current process's token.
        Dim returnLength As IntPtr
        If Not NativeApi.AdjustTokenPrivileges(hToken, False, newState, Marshal.SizeOf(prevState), prevState, returnLength) Then
            lastWin32Error = Marshal.GetLastWin32Error()
            Console.WriteLine("Error Adjusting Tokens: " & lastWin32Error.ToString)
        End If
    End Sub
#End Region
#Region "WriteMemory"
    Public Function ZeroMemory(ByVal Address As Long, ByVal length As Integer) As Boolean
        Dim buffer(length) As Byte
        For i = 0 To length
            buffer(i) = 0
        Next
        Return WriteMemory(Address, buffer)
    End Function
    Public Function WriteMemory(ByVal Address As Long, ByVal value As String) As Boolean
        Dim buffer As Byte() = ASCIIEncoding.UTF8.GetBytes(value)
        Return WriteMemory(Address, buffer)
    End Function
    Public Function WriteMemoryUni(ByVal Address As Long, ByVal value As String) As Boolean
        '           Dim str As System.Text.UnicodeEncoding = New System.Text.UnicodeEncoding

        '            Dim buffer As Byte() = str.GetBytes(value) 'ASCIIEncoding.Unicode.GetBytes(value & Chr(0) & Chr(0))
        Return WriteMemory(Address, Encoding.Unicode.GetBytes(value & Chr(0)))
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As UInt64) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Int64) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Double) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Single) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As SByte) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 1)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Byte) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 1)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Short) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 2)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As UShort) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 2)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As Integer) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 4)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal value As UInteger) As Boolean
        Dim buffer As Byte() = BitConverter.GetBytes(value)
        Return WriteMemory(Address, buffer, 4)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal buffer As Byte(), ByVal size As Integer) As Boolean
        Return NativeApi.WriteProcessMemory(Process_OpenedHandle, Address, buffer, size, IntPtr.Zero)
    End Function

    Public Function WriteMemory(ByVal Address As Long, ByVal buffer As Byte()) As Boolean
        Dim write_success As Boolean = NativeApi.WriteProcessMemory(Process_OpenedHandle, Address, buffer, buffer.Length, IntPtr.Zero)
        If Not write_success And Marshal.GetLastWin32Error() <> 299 Then
            If Marshal.GetLastWin32Error = 487 Then
                Console.WriteLine("Write Error: Invalid Address - " & Hex(Address))
            Else
                Console.WriteLine("Write Error: " & Marshal.GetLastWin32Error().ToString())
            End If

        End If
        Return write_success
    End Function
#End Region
#Region "ReadMemory"
    Public Function readPtr(ByVal address As Long, ByVal ParamArray offsets As Integer()) As Long
        Try
            'Return BitConverter.ToInt32(cMem.main.preader.ReadProcessMemory(address, 4, Nothing), 0)
            Dim res As Integer = readInt(address)
            Console.WriteLine(Hex(address))
            For Each i As Integer In offsets
                Console.WriteLine(Hex(res))
                res = readInt(res + i)
            Next
            Return res
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
        End Try
    End Function
    Public Function readInt(ByVal address As IntPtr) As Int32
        Try
            Return BitConverter.ToInt32(NativeApi.ReadMemory(Process_OpenedHandle, address, 4), 0)
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
        End Try
    End Function
    Public Function readLong(ByVal address As IntPtr) As Long
        Try
            Return BitConverter.ToInt32(NativeApi.ReadMemory(Process_OpenedHandle, address, 4), 0)
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
        End Try
    End Function
    Public Function readIntPtr(ByVal address As Long) As IntPtr
        Try
            Return BitConverter.ToUInt32(NativeApi.ReadMemory(Process_OpenedHandle, address, 4), 0)
        Catch ex As Exception
            Return -2
        End Try
    End Function
    Public Function readShort(ByVal address As Long) As Integer

        Return BitConverter.ToInt16(NativeApi.ReadMemory(Process_OpenedHandle, address, 2), 0)
    End Function
    Public Function readFloat(ByVal address As Long) As Single
        Return BitConverter.ToSingle(NativeApi.ReadMemory(Process_OpenedHandle, address, 4), 0)
    End Function
    Public Function readByte(ByVal address As Long) As Byte
        Return NativeApi.ReadMemory(Process_OpenedHandle, address, 2)(0)
    End Function
    Public Function readBytes(ByVal address As Long, ByVal size As Integer) As Integer
        Return BitConverter.ToInt16(NativeApi.ReadMemory(Process_OpenedHandle, address, size), 0)
    End Function
    Public Function readByteArray(ByVal address As Long, ByVal size As Integer) As Byte()
        Return NativeApi.ReadMemory(Process_OpenedHandle, address, size)
    End Function
    Public Function readString(ByVal address As Long, ByVal size As Long) As String
        Dim tstr As String = Encoding.ASCII.GetString(NativeApi.ReadMemory(Process_OpenedHandle, address, size))
        If tstr.Length > 0 Then
            If InStr(tstr, Chr(0) & Chr(0)) > 1 Then
                tstr = tstr.Substring(0, InStr(tstr, Chr(0) & Chr(0)) - 1)
            End If
        Else
            Return ""
        End If
        Return Replace(tstr, Chr(0), "")
    End Function
    Public Function readString2(ByVal address As Long, ByVal size As Long) As String
        Dim tstr As String = Encoding.ASCII.GetString(NativeApi.ReadMemory(Process_OpenedHandle, address, size))
        If tstr.Length > 0 Then
            If InStr(tstr, Chr(0)) > 1 Then
                tstr = tstr.Substring(0, InStr(tstr, Chr(0)) - 1)
            End If
        Else
            Return ""
        End If
        Return Replace(tstr, Chr(0), "")
    End Function
    Public Function readUniString(ByVal address As Long, ByVal size As Long) As String
        Dim tstr As String = Encoding.Unicode.GetString(NativeApi.ReadMemory(Process_OpenedHandle, address, size))
        If tstr.Length > 0 Then
            If InStr(tstr, Chr(0)) > 1 Then
                tstr = tstr.Substring(0, InStr(tstr, Chr(0)) - 1)
            End If
        Else
            Return ""
        End If
        Return Replace(tstr, Chr(0), "")
    End Function
#End Region
#Region "Memory Allocation/Free Memory"
    Public Function AllocMem(ByVal SizeOfAllocationInBytes As Integer) As Long
        Try
            Const MEM_COMMIT As Integer = &H1000
            Const PAGE_EXECUTE_READWRITE As Integer = &H40
            Dim pBlob As IntPtr = NativeApi.VirtualAllocEx(Process_Handle, Nothing, New IntPtr(SizeOfAllocationInBytes), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            If pBlob = IntPtr.Zero Then Throw New Exception
            Return pBlob
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try
    End Function
    Public Function VirtualFreeEx(ByVal address As IntPtr, ByVal size As IntPtr, ByVal dwtype As Integer) As Boolean
        NativeApi.VirtualFreeEx(address, address, size, &H4000)
    End Function
#End Region
#Region "Memory Protection"
    Public Sub RemoveProtection(ByVal AddressOfStart As Integer, ByVal SizeToRemoveProtectionInBytes As Integer)
        Try
            Const PAGE_EXECUTE_READWRITE As Integer = &H40
            Dim oldProtect As Integer
            If Not NativeApi.VirtualProtectEx(Process_Handle, New IntPtr(AddressOfStart), New IntPtr(SizeToRemoveProtectionInBytes), PAGE_EXECUTE_READWRITE, oldProtect) Then Throw New Exception
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub
#End Region
#Region "Injection"
    Public Function Inject(ByVal file As String) As Boolean
        Try
            Dim hKernel As Integer = NativeApi.LoadLibrary("kernel32.dll") 'Load the kernel so we can get some information about addresses for loadlibrarya
            Dim LoadLibraryAddr As Integer = NativeApi.GetProcAddress(hKernel, "LoadLibraryA") 'find the address of the function LoadLibraryA
            Dim StringMem As Integer = AllocMem(file.Length) 'Allocate some memory to write the file path
            WriteMemory(StringMem, file) 'Write the file to our allocated memory
            Dim ThreadID As Integer = 0 'Buffer for threadID returned from createremotethread
            Dim hThread As Integer = NativeApi.CreateRemoteThread(Process_Handle, IntPtr.Zero, 0, LoadLibraryAddr, New IntPtr(StringMem), 0, ThreadID) 'start a remote thread that calls the loadlibrary function
            NativeApi.WaitForSingleObject(hThread, Int32.MaxValue)
            Dim exit_code As Integer = 0
            NativeApi.GetExitCodeThread(hThread, exit_code)
            If Marshal.GetLastWin32Error() <> 0 Then
                Throw New Exception
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("Win32 Error: " & Marshal.GetLastWin32Error)
        End Try
    End Function
#End Region
#Region "Checks"
    Public Function isActiveWindow() As Boolean
        Try
            Dim title As New StringBuilder
            Dim hwnd As Integer = NativeApi.GetForegroundWindow
            If Process_Obj.MainWindowHandle = hwnd Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Shared Function GetForegroundTitle() As String
        Try
            Dim title As New StringBuilder
            Dim hwnd As Integer = NativeApi.GetForegroundWindow
            title.Capacity = NativeApi.GetWindowTextLength(hwnd) + 1
            NativeApi.GetWindowText(hwnd, title, title.Capacity)
            Return title.ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#End Region
#Region "Functions"

#End Region
End Class
