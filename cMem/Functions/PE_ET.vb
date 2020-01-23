Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Public Class PE_ET
    Public Structure IMAGE_DATA_DIRECTORY
        Public VirtualAddress As UInt32
        Public Size As UInt32
    End Structure
    Public Structure IMAGE_EXPORT_DIRECTORY
        Public Characteristics As UInt32
        Public TimeDateStamp As UInt32
        Public MajorVersion As UInt16
        Public MinorVersion As UInt16
        Public Name As UInt32
        Public Base As UInt32
        Public NumberOfFunctions As UInt32
        Public NumberOfNames As UInt32
        Public AddressOfFunctions As UInt32
        Public AddressOfNames As UInt32
        Public AddressOfNameOrdinals As UInt32
    End Structure
    Public Structure IMAGE_DOS_HEADER
        Public e_magic As UInt16
        Public e_cblp As UInt16
        Public e_cp As UInt16
        Public e_crlc As UInt16
        Public e_cparhdr As UInt16
        Public e_minalloc As UInt16
        Public e_maxalloc As UInt16
        Public e_ss As UInt16
        Public e_sp As UInt16
        Public e_csum As UInt16
        Public e_ip As UInt16
        Public e_cs As UInt16
        Public e_lfarlc As UInt16
        Public e_ovno As UInt16
        Public e_res_0 As UInt16
        Public e_res_1 As UInt16
        Public e_res_2 As UInt16
        Public e_res_3 As UInt16
        Public e_oemid As UInt16
        Public e_oeminfo As UInt16
        Public e_res2_0 As UInt16
        Public e_res2_1 As UInt16
        Public e_res2_2 As UInt16
        Public e_res2_3 As UInt16
        Public e_res2_4 As UInt16
        Public e_res2_5 As UInt16
        Public e_res2_6 As UInt16
        Public e_res2_7 As UInt16
        Public e_res2_8 As UInt16
        Public e_res2_9 As UInt16
        Public e_lfanew As UInt32
    End Structure
    Public Structure IMAGE_FILE_HEADER
        Public Machine As UInt16
        Public NumberOfSections As UInt16
        Public TimeDateStamp As UInt32
        Public PointerToSymbolTable As UInt32
        Public NumberOfSymbols As UInt32
        Public SizeOfOptionalHeader As UInt16
        Public Characteristics As UInt16
    End Structure

    <Flags()> Public Enum IFHMachine As UInt16
        x86 = &H14C
        Alpha = &H184
        ARM = &H1C0
        MIPS16R3000 = &H162
        MIPS16R4000 = &H166
        MIPS16R10000 = &H168
        PowerPCLE = &H1F0
        PowerPCBE = &H1F2
        Itanium = &H200
        MIPS16 = &H266
        Alpha64 = &H284
        MIPSFPU = &H366
        MIPSFPU16 = &H466
        x64 = &H8664
    End Enum

    <Flags()> Public Enum IFHCharacteristics As UInt16
        RelocationInformationStrippedFromFile = &H1
        Executable = &H2
        LineNumbersStripped = &H4
        SymbolTableStripped = &H8
        AggresiveTrimWorkingSet = &H10
        LargeAddressAware = &H20
        Supports16Bit = &H40
        ReservedBytesWo = &H80
        Supports32Bit = &H100
        DebugInfoStripped = &H200
        RunFromSwapIfInRemovableMedia = &H400
        RunFromSwapIfInNetworkMedia = &H800
        IsSytemFile = &H1000
        IsDLL = &H2000
        IsOnlyForSingleCoreProcessor = &H4000
        BytesOfWordReserved = &H8000
    End Enum
    Public Structure IMAGE_OPTIONAL_HEADER64
        Public Magic As UInt16
        Public MajorLinkerVersion As [Byte]
        Public MinorLinkerVersion As [Byte]
        Public SizeOfCode As UInt32
        Public SizeOfInitializedData As UInt32
        Public SizeOfUninitializedData As UInt32
        Public AddressOfEntryPoint As UInt32
        Public BaseOfCode As UInt32
        Public ImageBase As UInt64
        Public SectionAlignment As UInt32
        Public FileAlignment As UInt32
        Public MajorOperatingSystemVersion As UInt16
        Public MinorOperatingSystemVersion As UInt16
        Public MajorImageVersion As UInt16
        Public MinorImageVersion As UInt16
        Public MajorSubsystemVersion As UInt16
        Public MinorSubsystemVersion As UInt16
        Public Win32VersionValue As UInt32
        Public SizeOfImage As UInt32
        Public SizeOfHeaders As UInt32
        Public CheckSum As UInt32
        Public Subsystem As UInt16
        Public DllCharacteristics As UInt16
        Public SizeOfStackReserve As UInt64
        Public SizeOfStackCommit As UInt64
        Public SizeOfHeapReserve As UInt64
        Public SizeOfHeapCommit As UInt64
        Public LoaderFlags As UInt32
        Public NumberOfRvaAndSizes As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)> _
Public DataDirectory As IMAGE_DATA_DIRECTORY()
    End Structure
    Public Structure IMAGE_OPTIONAL_HEADER32
        Public Magic As UInt16
        Public MajorLinkerVersion As Byte
        Public MinorLinkerVersion As Byte
        Public SizeOfCode As UInt32
        Public SizeOfInitializedData As UInt32
        Public SizeOfUninitializedData As UInt32
        Public AddressOfEntryPoint As UInt32
        Public BaseOfCode As UInt32
        Public BaseOfData As UInt32
        Public ImageBase As UInt32
        Public SectionAlignment As UInt32
        Public FileAlignment As UInt32
        Public MajorOperatingSystemVersion As UInt16
        Public MinorOperatingSystemVersion As UInt16
        Public MajorImageVersion As UInt16
        Public MinorImageVersion As UInt16
        Public MajorSubsystemVersion As UInt16
        Public MinorSubsystemVersion As UInt16
        Public Win32VersionValue As UInt32
        Public SizeOfImage As UInt32
        Public SizeOfHeaders As UInt32
        Public CheckSum As UInt32
        Public Subsystem As UInt16
        Public DllCharacteristics As UInt16
        Public SizeOfStackReserve As UInt32
        Public SizeOfStackCommit As UInt32
        Public SizeOfHeapReserve As UInt32
        Public SizeOfHeapCommit As UInt32
        Public LoaderFlags As UInt32
        Public NumberOfRvaAndSizes As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)> _
Public DataDirectory As IMAGE_DATA_DIRECTORY()
    End Structure
    Const IMAGE_DOS_SIGNATURE = &H5A4D
    Const IMAGE_NT_SIGNATURE = &H4550
    Const IMAGE_NT_OPTIONAL_HDR64_MAGIC = &H20B
    Const IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B
    Const IMAGE_DIRECTORY_ENTRY_EXPORT = 0
    Public Shared Function RawDeserialize(ByVal rawData As Byte(), ByVal position As Integer, ByVal anyType As Type) As Object
        Dim rawsize As Integer = Marshal.SizeOf(anyType)
        If rawsize > rawData.Length Then
            Return Nothing
        End If
        Dim buffer As IntPtr = Marshal.AllocHGlobal(rawsize)
        Marshal.Copy(rawData, position, buffer, rawsize)
        Dim retobj As Object = Marshal.PtrToStructure(buffer, anyType)
        Marshal.FreeHGlobal(buffer)
        Return retobj
    End Function
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True)> _
    Public Shared Function GetProcAddress(ByVal hModule As IntPtr, ByVal procName As String) As UIntPtr
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)> _
 Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function
    Public Shared Function GetFunctionAddress(ByVal dll As String, ByVal function_name As String) As Long
        Return GetProcAddress(GetModuleHandle(dll), function_name).ToUInt32
    End Function
    Public Shared Function GetFunctionSize(ByVal function_name As String) As Long
        ' cMem.main.sreader = New ProcessMemoryReader
        ' cMem.main.setSelf()
        Dim Is64Bit As Boolean = False
        Dim ExportFunctionTableVA As Long = 0
        Dim ExportNameTableVA As Long = 0
        Dim ExportOrdinalTableVA As Long = 0
        Dim dh As New IMAGE_DOS_HEADER
        Dim fh As New IMAGE_FILE_HEADER
        Dim dd As New IMAGE_DATA_DIRECTORY
        Dim ExportTable As New IMAGE_EXPORT_DIRECTORY
        Dim OptHeader64 As New IMAGE_OPTIONAL_HEADER64
        Dim OptHeader32 As New IMAGE_OPTIONAL_HEADER32
        Dim signature As Integer = 0
        Dim pm As ProcessModule = LocalProcess.ModuleObj("kernel32.dll")
        Dim buffer() As Byte
        Dim p_objTarget As IntPtr
        p_objTarget = Marshal.AllocHGlobal(Marshal.SizeOf(dh))
        ' dh = NativeApi.ReadMemoryobj(pm.BaseAddress, Marshal.SizeOf(dh), 0)
        buffer = NativeApi.ReadMemory(pm.BaseAddress, Marshal.SizeOf(dh), 0)
        dh = RawDeserialize(buffer, 0, GetType(IMAGE_DOS_HEADER))
        If dh.e_magic <> IMAGE_DOS_SIGNATURE Then
            MsgBox("Failed to get image dos header")
            Exit Function
        End If

        signature = BitConverter.ToInt32(NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew, 4, Nothing), 0)

        If signature <> IMAGE_NT_SIGNATURE Then
            MsgBox("Failed to get nt image signature")
            Exit Function
        End If
        buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4, Marshal.SizeOf(fh), 0)
        fh = RawDeserialize(buffer, 0, GetType(IMAGE_FILE_HEADER))
        If (fh.SizeOfOptionalHeader = Marshal.SizeOf(OptHeader64)) Then
            Is64Bit = True
        ElseIf (fh.SizeOfOptionalHeader = Marshal.SizeOf(OptHeader32)) Then
            Is64Bit = False
        Else
            MsgBox("Failed to detect optional header size")
            Exit Function
        End If

        If Is64Bit Then
            buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4 + Marshal.SizeOf(fh), Marshal.SizeOf(OptHeader64), 0)
            OptHeader64 = RawDeserialize(buffer, 0, GetType(IMAGE_OPTIONAL_HEADER64))
            If (OptHeader64.Magic <> IMAGE_NT_OPTIONAL_HDR64_MAGIC) Then
                MsgBox("Optional 64 bit header failed to match magic number")
                Exit Function
            End If
        Else
            buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4 + Marshal.SizeOf(fh), Marshal.SizeOf(OptHeader32), 0)
            OptHeader32 = RawDeserialize(buffer, 0, GetType(IMAGE_OPTIONAL_HEADER32))
            If (OptHeader32.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC) Then
                Console.WriteLine("optional 32 magic number: " & Hex(OptHeader32.Magic) & " != " & Hex(IMAGE_NT_OPTIONAL_HDR32_MAGIC))
                MsgBox("Optional 32 bit header failed to match magic number")
                Exit Function
            End If
        End If

        If (Is64Bit And OptHeader64.NumberOfRvaAndSizes >= IMAGE_DIRECTORY_ENTRY_EXPORT + 1) Then
            dd.VirtualAddress = OptHeader64.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
            dd.Size = OptHeader64.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).Size
        ElseIf (OptHeader32.NumberOfRvaAndSizes >= IMAGE_DIRECTORY_ENTRY_EXPORT + 1) Then
            dd.VirtualAddress = OptHeader32.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
            dd.Size = OptHeader32.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).Size
        Else
            MsgBox("Data Directory failed")
            Exit Function
        End If
        buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dd.VirtualAddress, Marshal.SizeOf(ExportTable), 0)
        ExportTable = RawDeserialize(buffer, 0, GetType(IMAGE_EXPORT_DIRECTORY))

        ExportFunctionTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfFunctions
        ExportNameTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfNames
        ExportOrdinalTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfNameOrdinals
        Dim ExportFunctionTable As Long()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfFunctions)
        Dim ExportNameTable As Long()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfNames)
        Dim ExportOrdinalTable As Short()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfNames)
        buffer = NativeApi.ReadMemory(ExportFunctionTableVA, ExportTable.NumberOfFunctions * 4, 0)
        ExportFunctionTable = ByteArrayToIntegerArray(buffer)
        buffer = NativeApi.ReadMemory(ExportNameTableVA, ExportTable.NumberOfNames * 4, 0)
        ExportNameTable = ByteArrayToIntegerArray(buffer)
        buffer = NativeApi.ReadMemory(ExportOrdinalTableVA, ExportTable.NumberOfNames * 2, 0)
        ExportOrdinalTable = ByteArrayToShortArray(buffer)

        ' ExportFunctionTable = RawDeserialize(buffer, 0, GetType(Array))
        Dim i As Integer = 0
        Dim length As Long = Long.MaxValue
        Dim c_address As Long = GetProcAddress(GetModuleHandle("Kernel32.dll"), function_name).ToUInt32
        For i = 0 To ExportNameTable.Length - 1
            Dim str As String = readString(pm.BaseAddress.ToInt32 + ExportNameTable(i), 128)
            ' If LCase(str) = LCase(function_name) Then
            ' Console.WriteLine("Made it!")
            ' str = readString(pm.BaseAddress.ToInt32 + ExportNameTable(i + 1), 128)
            ' Return GetProcAddress(GetModuleHandle("Kernel32.dll"), str).ToUInt32 - c_address
            ' End If
            length = GetProcAddress(GetModuleHandle("Kernel32.dll"), str).ToUInt32 - c_address
            If length > 0 Then
                Return length
            End If
            'If diff > 0 And length > diff Then
            ' length = diff
            ' Console.WriteLine("Str: " & str & "Address: " & Hex(GetProcAddress(GetModuleHandle("Kernel32.dll"), str).ToUInt32))
            ' End If
        Next
        Return 0
    End Function
    Public Shared Sub EnumModule()
        '  cMem.main.sreader = New ProcessMemoryReader
        '  cMem.main.setSelf()
        Dim Is64Bit As Boolean = False
        Dim ExportFunctionTableVA As Long = 0
        Dim ExportNameTableVA As Long = 0
        Dim ExportOrdinalTableVA As Long = 0
        Dim dh As New IMAGE_DOS_HEADER
        Dim fh As New IMAGE_FILE_HEADER
        Dim dd As New IMAGE_DATA_DIRECTORY
        Dim ExportTable As New IMAGE_EXPORT_DIRECTORY
        Dim OptHeader64 As New IMAGE_OPTIONAL_HEADER64
        Dim OptHeader32 As New IMAGE_OPTIONAL_HEADER32
        Dim signature As Integer = 0
        Dim pm As ProcessModule = LocalProcess.ModuleObj("kernel32.dll")
        Dim buffer() As Byte
        Dim p_objTarget As IntPtr
        p_objTarget = Marshal.AllocHGlobal(Marshal.SizeOf(dh))
        ' dh = NativeApi.ReadMemoryobj(pm.BaseAddress, Marshal.SizeOf(dh), 0)
        buffer = NativeApi.ReadMemory(pm.BaseAddress, Marshal.SizeOf(dh), 0)
        dh = RawDeserialize(buffer, 0, GetType(IMAGE_DOS_HEADER))
        If dh.e_magic <> IMAGE_DOS_SIGNATURE Then
            MsgBox("Failed to get image dos header")
            Exit Sub
        End If
        Console.WriteLine("E_LFANEW " & dh.e_lfanew)
        Console.WriteLine("E_MAGIC " & dh.e_magic)

        signature = BitConverter.ToInt32(NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew, 4, Nothing), 0)
        Console.WriteLine("Signature " & signature)
        If signature <> IMAGE_NT_SIGNATURE Then
            MsgBox("Failed to get nt image signature")
            Exit Sub
        End If
        buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4, Marshal.SizeOf(fh), 0)
        fh = RawDeserialize(buffer, 0, GetType(IMAGE_FILE_HEADER))
        If (fh.SizeOfOptionalHeader = Marshal.SizeOf(OptHeader64)) Then
            Is64Bit = True
        ElseIf (fh.SizeOfOptionalHeader = Marshal.SizeOf(OptHeader32)) Then
            Is64Bit = False
        Else
            MsgBox("Failed to detect optional header size")
            Exit Sub
        End If

        If Is64Bit Then
            buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4 + Marshal.SizeOf(fh), Marshal.SizeOf(OptHeader64), 0)
            OptHeader64 = RawDeserialize(buffer, 0, GetType(IMAGE_OPTIONAL_HEADER64))
            If (OptHeader64.Magic <> IMAGE_NT_OPTIONAL_HDR64_MAGIC) Then
                MsgBox("Optional 64 bit header failed to match magic number")
                Exit Sub
            End If
        Else
            buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dh.e_lfanew + 4 + Marshal.SizeOf(fh), Marshal.SizeOf(OptHeader32), 0)
            OptHeader32 = RawDeserialize(buffer, 0, GetType(IMAGE_OPTIONAL_HEADER32))
            If (OptHeader32.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC) Then
                Console.WriteLine("optional 32 magic number: " & Hex(OptHeader32.Magic) & " != " & Hex(IMAGE_NT_OPTIONAL_HDR32_MAGIC))
                MsgBox("Optional 32 bit header failed to match magic number")
                Exit Sub
            End If
        End If

        If (Is64Bit And OptHeader64.NumberOfRvaAndSizes >= IMAGE_DIRECTORY_ENTRY_EXPORT + 1) Then
            dd.VirtualAddress = OptHeader64.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
            dd.Size = OptHeader64.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).Size
        ElseIf (OptHeader32.NumberOfRvaAndSizes >= IMAGE_DIRECTORY_ENTRY_EXPORT + 1) Then
            dd.VirtualAddress = OptHeader32.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
            dd.Size = OptHeader32.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).Size
        Else
            MsgBox("Data Directory failed")
            Exit Sub
        End If
        buffer = NativeApi.ReadMemory(pm.BaseAddress.ToInt32 + dd.VirtualAddress, Marshal.SizeOf(ExportTable), 0)
        ExportTable = RawDeserialize(buffer, 0, GetType(IMAGE_EXPORT_DIRECTORY))

        ExportFunctionTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfFunctions
        ExportNameTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfNames
        ExportOrdinalTableVA = pm.BaseAddress.ToInt32 + ExportTable.AddressOfNameOrdinals
        Dim ExportFunctionTable As Long()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfFunctions)
        Dim ExportNameTable As Long()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfNames)
        Dim ExportOrdinalTable As Short()
        ReDim Preserve ExportFunctionTable(ExportTable.NumberOfNames)
        buffer = NativeApi.ReadMemory(ExportFunctionTableVA, ExportTable.NumberOfFunctions * 4, 0)
        ExportFunctionTable = ByteArrayToIntegerArray(buffer)
        buffer = NativeApi.ReadMemory(ExportNameTableVA, ExportTable.NumberOfNames * 4, 0)
        ExportNameTable = ByteArrayToIntegerArray(buffer)
        buffer = NativeApi.ReadMemory(ExportOrdinalTableVA, ExportTable.NumberOfNames * 2, 0)
        ExportOrdinalTable = ByteArrayToShortArray(buffer)

        ' ExportFunctionTable = RawDeserialize(buffer, 0, GetType(Array))
        Console.WriteLine("Exported Function Count: " & ExportTable.NumberOfFunctions)
        Console.WriteLine("Exported Name Count: " & ExportNameTable.Length)
        Dim i As Integer = 0
        For i = 0 To ExportNameTable.Length - 1
            Dim str As String = readString(pm.BaseAddress.ToInt32 + ExportNameTable(i), 128)
            Console.WriteLine(str & " Address?: " & Hex(GetProcAddress(GetModuleHandle("Kernel32.dll"), str).ToUInt32))
        Next
        '     For i = 0 To ExportFunctionTable.Length - 1
        ' Dim str As String = readString(pm.BaseAddress.ToInt32 + ExportFunctionTable(i), 128)
        ' Console.WriteLine(str)
        '            If InStr(str, "Process") Then
        ' Console.WriteLine(str)
        ' End If
        ' If Left(str, 1) = "#" Then 'ordinal
        ' Dim ordinal As Integer = 0
        '     For i=0
        ' End If
        'If InStr(str, "NTDLL") Then
        ' Console.WriteLine(str)
        ' End If
        'Next
    End Sub
    Private Shared Function ByteArrayToShortArray(ByVal bytearray As Byte()) As Short()
        Dim i As Long = 0
        Dim l As Short()
        ReDim Preserve l(bytearray.Length / 2 - 1)
        Dim c As Long = 0
        For i = 0 To bytearray.Length
            Dim ba As Byte() = {bytearray(i), bytearray(i + 1)}
            l(c) = BitConverter.ToInt16(ba, 0)
            c = c + 1
            i = i + 1
            If i = bytearray.Length - 1 Then Exit For
        Next
        Return l
    End Function
    Private Shared Function ByteArrayToIntegerArray(ByVal bytearray As Byte()) As Long()
        Dim i As Long = 0
        Dim l As Long()
        ReDim Preserve l(bytearray.Length / 4 - 1)
        Dim c As Long = 0
        For i = 0 To bytearray.Length
            Dim ba As Byte() = {bytearray(i), bytearray(i + 1), bytearray(i + 2), bytearray(i + 3)}
            l(c) = BitConverter.ToInt32(ba, 0)
            c = c + 1
            i = i + 3
            If i = bytearray.Length - 1 Then Exit For
        Next
        Return l
    End Function
    Private Shared Function readString(ByVal address As Long, ByVal size As Long) As String
        Dim tstr As String = Encoding.ASCII.GetString(NativeApi.ReadMemory(address, size, Nothing))
        If InStr(tstr, Chr(0)) Then
            tstr = Left(tstr, InStr(tstr, Chr(0)) - 1)
        End If
        Return tstr
    End Function
End Class
