# cMem
 .Net Memory Library


Library written by iam_clint | E-mail : clint.ban@gmail.com , very useful for memory management / Injection, among other things ...

https://github.com/devoyster/IntXLib

Sample source using it . 

```vb
'This code was written by clint.ban@gmail.com
'You may use this code for learning purposes only.
Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Public Class Form1
    Public Patch1 As Integer = &H531C8B 'Have to bypass the if statement -- if (v10) since I want to make it lie about connected clients even if its 0
    Public PushPatch As Integer = &H531C91 'This is where it pushes the pointer to the string onto the stack to make the function call (c++ sprintf)
    Public ServerName As String = &H185C424 'Pointer to the location of the server name
    Public data As String = "12"
    Public data_address As Integer
    Public ServerName_Address As Integer
    Public current_name As Integer = 0
    Public mem As New cMem.Memory
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim proclist As New cMem.ProcessList("iw3mp") 'Get a list of processes with iw3mp
            For Each item As Process In proclist.Processes 'iterate the processes
                If item.MainWindowTitle = "CoD4 Console" Then 'The title of the server is cod4 console
                    mem.Attach(item)
                End If
            Next
            If mem.Process_Handle = 0 Then
                MsgBox("Problem Finding Cod4") 'couldn't find the server
                Application.Exit()
            End If
            ServerName_Address = mem.readPtr(ServerName, 12) 'The server name is read from a pointer ->(servername+0xc) this function reads the  integer value of 0x185C424 then it adds 12 (0xc) and reads its value then reads the value of that as the final address
 
            Dim patch_val As Byte() = {&H90, &H90} '0x90=NOP (No Operation) takes the if statement which is 2 bytes 0x85 0xF6 orginal hex (test esi, esi) as assembly and it becomes (nop, nop) <-- does absolutely nothing
            mem.WriteMemory(Patch1, patch_val) 'write this patch to the if statement
            Dim buffer As Byte() = ASCIIEncoding.UTF8.GetBytes(data) 'Convert our string to a byte array "12"
            Dim address As Integer = mem.AllocMem(buffer.Count + 1) 'Allocate a memory region the size+1 of our byte array from above
            data_address = address 'The returned value from allocmem is the address at which it created the memory range
            mem.WriteMemory(address, data) 'write our data into the allocated memory
            Dim bytes() As Byte = BitConverter.GetBytes(address) 'get a byte array of the address for the allocated memory so we can write the pointer in hex
            mem.WriteMemory(PushPatch, bytes) 'write the patch to memory so now instead of its original push of 0c 7c 6c 00 it will be a push with the pointer of our allocated memory with the string in it
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
    End Sub
 
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text.Length = 1 Then TextBox1.Text &= " "
        mem.WriteMemory(data_address, TextBox1.Text)
    End Sub
 
    Private Sub CheckBox1_CheckedChanged(ByVal sender As CheckBox, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If sender.Checked Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub
 
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim rnd As New Random()
        Dim val As Integer = rnd.Next(10, 24)
        mem.WriteMemory(data_address, val.ToString())
    End Sub
 
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        current_name += 1
        Dim data() As String = TextBox2.Text.Split(vbCrLf)
        If current_name >= data.Length Then current_name = 0
        mem.ZeroMemory(ServerName_Address, 40)
        mem.WriteMemory(ServerName_Address, data(current_name).Replace(vbCr, "").Replace(vbLf, ""))
    End Sub
End Class
```

a simple injector code using it

```vb
'This code was written by clint.ban@gmail.com
'You may use this code for learning purposes only.
Imports cMem
Imports System.Runtime.InteropServices
Public Class Form1
    Dim mem As New cMem.Memory
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        refreshProcessList()
    End Sub
    Sub refreshProcessList()
        Dim pl As New cMem.ProcessList()
        ProcessList.Items.Clear()
        For Each p As Process In pl.Processes.Values
            Try
                Dim str(3) As String
                str(0) = p.ProcessName
                str(1) = p.MainWindowTitle
                str(2) = p.Id
                Dim item As New ListViewItem(str)
                ProcessList.Items.Add(item)
            Catch ex As Exception
            End Try
        Next
    End Sub
 
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'refreshProcessList()
    End Sub
 
    Private Sub DllBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DllBrowse.Click
        Dim ld As OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        With ld
            .DefaultExt = ".dll"
            .Filter = "Dll Files|*.dll"
            .InitialDirectory = Application.StartupPath & "\Waypoints"
        End With
        If ld.ShowDialog = Windows.Forms.DialogResult.OK Then
            DllPath.Text = ld.FileName
        End If
    End Sub
 
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim handle As Integer = Integer.Parse(ProcessList.SelectedItems(0).SubItems(2).Text.ToString)
        Try
            Dim pl As New cMem.ProcessList()
            For Each p As Process In pl.Processes.Values
                If handle = p.Id Then
                    mem.Attach(p)
                End If
            Next
            mem.Inject(DllPath.Text)
            If Marshal.GetLastWin32Error() <> 0 Then
                MsgBox(Marshal.GetLastWin32Error())
            End If
        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
```

