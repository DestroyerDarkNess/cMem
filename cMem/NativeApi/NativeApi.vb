Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Public Class NativeApi
    Public Delegate Function keyboardHookProc(ByVal code As Integer, ByVal wParam As Integer, ByRef lParam As keyboardHookStruct) As Integer

    <DllImport("advapi32.dll", SetLastError:=True)> _
     Public Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function
    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function LookupPrivilegeValue(ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Boolean
    End Function
    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLengthInBytes As Integer, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLengthInBytes As Integer) As Boolean
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function VirtualFreeEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal dwSize As IntPtr, ByVal FreeType As Integer) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function VirtualAllocEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal dwSize As IntPtr, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function VirtualProtectEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal dwSize As IntPtr, ByVal newProtect As Integer, ByRef oldProtect As Integer) As Boolean
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function WriteProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <[In](), Out()> ByVal buffer() As Byte, ByVal size As UInt32, ByRef lpNumberOfBytesWritten As IntPtr) As Int32
    End Function
    <DllImport("kernel32.dll")> _
    Public Shared Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <[In](), Out()> ByVal buffer As Byte(), ByVal size As UInt32, ByRef lpNumberOfBytesRead As IntPtr) As Int32
    End Function
    Public Shared Function ReadMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, ByVal size As UInt32) As Byte()
        Dim buffer As Byte() = New Byte(size - 1) {}
        Dim ptrBytesRead As IntPtr
        ReadProcessMemory(hProcess, lpBaseAddress, buffer, size, ptrBytesRead)
        Return buffer
    End Function

    <DllImport("kernel32", SetLastError:=True)> _
Public Shared Function GetExitCodeThread( _
ByVal handle As IntPtr, _
ByRef ExitCode As UInt32) As UInt32
    End Function
    <DllImport("kernel32", SetLastError:=True)> _
    Public Shared Function WaitForSingleObject( _
    ByVal handle As IntPtr, _
    ByVal milliseconds As UInt32) As UInt32
    End Function
    <DllImport("kernel32")> _
     Public Shared Function LoadLibrary(ByVal lpLibFileName As String) As Integer
    End Function
    <DllImport("kernel32")> _
    Public Shared Function FreeLibrary(ByVal hLibModule As Integer) As Boolean
    End Function
    <DllImport("kernel32")> _
    Public Shared Function CreateRemoteThread(ByVal hProcess As IntPtr, ByVal lpThreadAttributes As IntPtr, ByVal dwStackSize As UInt32, ByVal lpStartAddress As IntPtr, ByVal lpParameter As IntPtr, ByVal dwCreationFlags As IntPtr, ByRef lpThreadId As UInt32) As IntPtr
    End Function
    <DllImport("kernel32", CharSet:=CharSet.Ansi)> _
    Public Shared Function GetProcAddress(ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    End Function
    <DllImport("ntdll.dll")> _
    Public Shared Function ZwReadVirtualMemory(ByVal hProcess As IntPtr, ByVal address As IntPtr, <Out()> ByRef buffer As Byte(), ByVal Size As UInt64, ByRef ReadSize As UInt64) As UInt32
    End Function
    <DllImport("kernel32.dll")> _
    Public Shared Function OpenProcess(ByVal dwDesiredAccess As UInt32, ByVal bInheritHandle As Int32, ByVal dwProcessId As UInt32) As IntPtr
    End Function
    <DllImport("kernel32.dll")> _
    Public Shared Function RtlMoveMemory(ByVal pDst As IntPtr, ByVal pSrc As String, ByVal ByteLen As Long) As IntPtr
    End Function
    <DllImport("User32.dll")> _
   Public Shared Function GetWindowThreadProcessId(ByVal hWnd As Int32, ByRef lpdwProcessId As Int32) As UInt32
    End Function
    <DllImport("kernel32.dll")> _
    Public Shared Function OpenThread(ByVal dwDesiredAccess As UInteger, ByVal bInheritHandle As Boolean, ByVal dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("kernel32.dll")> _
        Public Shared Function TerminateThread(ByVal hThread As IntPtr, ByVal dwExitCode As UInteger) As Boolean
    End Function
    <DllImport("kernel32.dll")> _
        Public Shared Function SuspendThread(ByVal hThread As IntPtr) As UInteger
    End Function
    <DllImport("kernel32.dll")> _
    Public Shared Function VirtualQueryEx(ByVal hProcess As Integer, ByVal lpAddress As Integer, <MarshalAs(UnmanagedType.Struct)> ByRef lpBuffer As Mbi.MEMORY_BASIC_INFORMATION, ByVal dwLength As Integer) As Integer
    End Function


    <DllImport("kernel32.dll")> _
        Public Shared Function ResumeThread(ByVal hThread As IntPtr) As UInteger
    End Function
    <DllImport("kernel32.dll")> _
        Public Shared Function CloseHandle(ByVal hObject As IntPtr) As Boolean
    End Function
    <DllImport("ntdll.dll")> _
        Public Shared Function NtQueryInformationThread(ByVal handle As IntPtr, ByVal type As UInteger, ByRef return_val As THREAD_BASIC_INFORMATION, ByVal length As UInteger, ByVal bytes_read As IntPtr) As UInteger
    End Function
    <DllImport("ntdll.dll")> _
        Public Shared Function NtQueryInformationThread(ByVal handle As IntPtr, ByVal type As UInteger, ByRef return_val As Long, ByVal length As UInteger, ByVal bytes_read As IntPtr) As UInteger
    End Function
    <DllImport("kernel32.dll")> _
        Public Shared Sub GetNativeSystemInfo(ByRef lpSystemInfo As SYSTEM_INFO)
    End Sub
    <DllImport("kernel32.dll")> _
    Public Shared Function GetModuleHandleA(ByVal lpModuleName As String) As Long
    End Function
    <DllImport("kernel32.dll")> _
        Public Shared Sub GetSystemInfo(ByRef lpSystemInfo As SYSTEM_INFO)
    End Sub
    <DllImport("User32.dll")> _
        Public Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndParent As IntPtr) As IntPtr
    End Function
    <DllImport("User32.dll")> _
        Public Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("User32.dll")> _
        Public Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpWindowText As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("User32.dll")> _
        Public Shared Function SetForegroundWindow(ByVal hWnd As Integer) As Int32
    End Function
    <DllImport("User32.dll")> _
        Public Shared Function VkKeyScan(ByVal ch As Char) As Short
    End Function
    <DllImport("User32.dll")> _
        Public Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal callback As keyboardHookProc, ByVal hInstance As IntPtr, ByVal threadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function UnhookWindowsHookEx(ByVal hInstance As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")> _
    Public Shared Function CallNextHookEx(ByVal idHook As IntPtr, ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As keyboardHookStruct) As Integer
    End Function
    <DllImport("user32.dll")> _
    Public Shared Function CallNextHookEx(ByVal idHook As IntPtr, ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MouseHookStruct) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As keyboardHookStruct) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("User32.dll", SetLastError:=False, CallingConvention:=CallingConvention.StdCall, _
          CharSet:=CharSet.Auto)> _
    Public Shared Function MapVirtualKey(ByVal uCode As UInt32, ByVal uMapType As MapVirtualKeyMapTypes) As UInt32
    End Function
    Public Enum MapVirtualKeyMapTypes As UInt32
        ''' <summary>uCode is a virtual-key code and is translated into a scan code.
        ''' If it is a virtual-key code that does not distinguish between left- and
        ''' right-hand keys, the left-hand scan code is returned.
        ''' If there is no translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_VSC = &H0

        ''' <summary>uCode is a scan code and is translated into a virtual-key code that
        ''' does not distinguish between left- and right-hand keys. If there is no
        ''' translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VSC_TO_VK = &H1

        ''' <summary>uCode is a virtual-key code and is translated into an unshifted
        ''' character value in the low-order word of the return value. Dead keys (diacritics)
        ''' are indicated by setting the top bit of the return value. If there is no
        ''' translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_CHAR = &H2

        ''' <summary>Windows NT/2000/XP: uCode is a scan code and is translated into a
        ''' virtual-key code that distinguishes between left- and right-hand keys. If
        ''' there is no translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VSC_TO_VK_EX = &H3

        ''' <summary>Not currently documented
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_VSC_EX = &H4
    End Enum
    Public Structure THREAD_BASIC_INFORMATION
        Public ExitStatus As Boolean
        Public TebBaseAddress As IntPtr
        Public processid As UInteger
        Public threadid As UInteger
        Public AffinityMask As UInteger
        Public Priority As UInteger
        Public BasePriority As UInteger
    End Structure
    Public Structure SYSTEM_INFO
        Public wProcessorArchitecture As UShort
        Public wReserved As UShort
        Public dwPageSize As UInteger
        Public lpMinimumApplicationAddress As IntPtr
        Public lpMaximumApplicationAddress As IntPtr
        Public dwActiveProcessorMask As UIntPtr
        Public dwNumberOfProcessors As UInteger
        Public dwProcessorType As UInteger
        Public dwAllocationGranularity As UInteger
        Public wProcessorLevel As UShort
        Public wProcessorRevision As UShort
    End Structure
    Public Enum Platform
        X86
        X64
        Unknown
    End Enum
    Public Shared length_mismatch As Long = &HC0000004
    Public Enum Processor_Architecture As UShort
        INTEL = 0
        IA64 = 6
        AMD64 = 9
        UNKNOWN = &HFFFF
    End Enum
    Public Enum Rights As Integer
        STANDARD_RIGHTS_REQUIRED = &HF0000
        STANDARD_RIGHTS_READ = &H20000
        STANDARD_RIGHTS_WRITE = &H20000
        STANDARD_RIGHTS_EXECUTE = &H20000
        STANDARD_RIGHTS_ALL = &H1F0000
        SPECIFIC_RIGHTS_ALL = &HFFFF
    End Enum
    Public Enum Tokens As Integer
        ASSIGN_PRIMARY = &H1
        DUPLICATE = &H2
        IMPERSONATE = &H4
        QUERY = &H8
        QUERY_SOURCE = &H10
        ADJUST_PRIVILEGES = &H20
        ADJUST_GROUPS = &H40
        ADJUST_DEFAULT = &H80
        ADJUST_SESSIONID = &H100
    End Enum
    Public Shared ANYSIZE_ARRAY As Integer = 1
    Public Shared SE_PRIVILEGE_ENABLED As Integer = &H2
    <StructLayout(LayoutKind.Sequential)> _
  Public Structure LUID
        Public LowPart As UInt32
        Public HighPart As UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure LUID_AND_ATTRIBUTES
        Public Luid As LUID
        Public Attributes As UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)> _
        Public Privileges() As LUID_AND_ATTRIBUTES
    End Structure
    <Flags()> _
Public Enum ProcessAccessType
        PROCESS_TERMINATE = (&H1)
        PROCESS_CREATE_THREAD = (&H2)
        PROCESS_SET_SESSIONID = (&H4)
        PROCESS_VM_OPERATION = (&H8)
        PROCESS_VM_READ = (&H10)
        PROCESS_VM_WRITE = (&H20)
        PROCESS_DUP_HANDLE = (&H40)
        PROCESS_CREATE_PROCESS = (&H80)
        PROCESS_SET_QUOTA = (&H100)
        PROCESS_SET_INFORMATION = (&H200)
        PROCESS_QUERY_INFORMATION = (&H400)
    End Enum
    Public Structure keyboardHookStruct
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure
    'Point structure declaration.
    <StructLayout(LayoutKind.Sequential)> Public Structure Point
        Public x As Integer
        Public y As Integer
    End Structure

    'MouseHookStruct structure declaration.
    <StructLayout(LayoutKind.Sequential)> Public Structure MouseHookStruct
        Public pt As Point
        Public hwnd As Integer
        Public wHitTestCode As Integer
        Public dwExtraInfo As Integer
    End Structure

    Public Shared PROCESSOR_ARCHITECTURE_INTEL As UShort = 0
    Public Shared PROCESSOR_ARCHITECTURE_IA64 As UShort = 6
    Public Shared PROCESSOR_ARCHITECTURE_AMD64 As UShort = 9
    Public Shared PROCESSOR_ARCHITECTURE_UNKNOWN As UShort = &HFFFF
End Class
