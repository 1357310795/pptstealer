Imports System.ComponentModel
Imports System.Diagnostics
Imports IWshRuntimeLibrary

Public Class Form1
    Public exe_paths As List(Of String)
    Public ext_list As List(Of String)
    Public copy_late_time As Int32
    Public ok_list As List(Of Int32)
    Public task_list As List(Of Process)
    Public timer_list As List(Of Timer)
    Public real_close As Boolean
    Public main_timer As Timer

    Public ti As Timer
    Public slidetime As Int32


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox(GetTypeFilePath("txt"))
        'MsgBox(GetDriveType(System.IO.Path.GetPathRoot("C:\windows\system32\notepad.exe")))
        'MsgBox(My.Computer.FileSystem.GetFileInfo("C:\windows\system32\notepad.exe").Length)
        'MsgBox(My.Computer.FileSystem.GetFileInfo("C:\windows\system32\notepad.exe").Extension)
        exe_paths = New List(Of String)
        ext_list = New List(Of String)
        task_list = New List(Of Process)
        timer_list = New List(Of Timer)
        disk_save_paths = New List(Of String)
        ok_list = New List(Of Integer)

        read_settings()
        ''''
        'ext_list.Add("xlsx")
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
        Me.Width = Panel1.Width

        NotifyIcon1.Icon = Me.Icon
        NotifyIcon1.Text = "cs"
        NotifyIcon1.Visible = True

        main_timer = New Timer
        main_timer.Interval = 5000
        AddHandler main_timer.Tick, AddressOf main_timer_Tick

        Dim mini_timer = New Timer
        mini_timer.Interval = 1000
        AddHandler mini_timer.Tick, AddressOf mini_timer_Tick
        mini_timer.Start()

        CheckBox1.Checked = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            main_timer.Start()
        Else
            main_timer.Stop()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        For Each pro As Process In Process.GetProcesses
            'System.IO.Path.GetPathRoot()
            Console.WriteLine(GetCommandLine(pro))
        Next
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(0, "111", "222", ToolTipIcon.Info)
    End Sub

    Private Sub TypeCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged,
            CheckBox3.CheckedChanged,
            CheckBox4.CheckedChanged,
            CheckBox5.CheckedChanged,
            CheckBox6.CheckedChanged,
            CheckBox7.CheckedChanged,
            CheckBox8.CheckedChanged


        Dim ext As String = sender.text
        Dim open_path As String = GetTypeFilePath(ext)
        If open_path = "" Then
            writetext("未找到扩展名" & ext & "的打开程序！")
            Exit Sub
        End If
        Dim exe_path As String = calc_exe_path(open_path)
        Console.WriteLine(exe_path)
        writetext(exe_path)
        If sender.checked Then
            exe_paths.Add(exe_path)
            ext_list.Add("." & sender.text.Tolower)
        Else
            exe_paths.Remove(exe_path)
            ext_list.Remove("." & sender.text.Tolower)
        End If
    End Sub
    Private Sub writetext(ByVal txt As String)
        Dim tmp As List(Of String)
        tmp = RichTextBox1.Lines.ToList
        tmp.Add(txt)
        RichTextBox1.Lines = tmp.ToArray
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not real_close Then
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
        End If
    End Sub


    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        End If
    End Sub

    Private Sub ContextMenuStrip1_Click(sender As Object, e As EventArgs) Handles ContextMenuStrip1.Click
        real_close = True
        Me.Close()
    End Sub

    Private Sub NotifyIcon1_DoubleClick(sender As Object, e As EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub main_timer_Tick(sender As Object, e As EventArgs)
        'sender.stop()
        Dim tmp As List(Of Process) = New List(Of Process)
        For Each p As Process In Process.GetProcesses
            Dim flag As Boolean
            For Each i As String In exe_paths
                If i.ToLower.IndexOf(p.ProcessName.ToLower) <> -1 Then
                    If GetCommandLine(p).ToLower.IndexOf(i.ToLower) <> -1 Then
                        flag = True
                        Exit For
                    End If
                End If
            Next
            If flag Then
                tmp.Add(p)
                flag = False
            End If
            'Console.WriteLine(p.ProcessName)
        Next
        Dim now_pid_list As New List(Of Int32)
        For Each p In tmp
            If ok_list.IndexOf(p.Id) = -1 Then
                task_list.Add(p)
                ok_list.Add(p.Id)
                Dim newt As New Timer
                newt.Interval = NumericUpDown1.Value * 1000
                newt.Tag = task_list.Count - 1
                AddHandler newt.Tick, AddressOf copy_timer_Tick
                timer_list.Add(newt)
                newt.Start()
            End If
            now_pid_list.Add(p.Id)
        Next

        For i = ok_list.Count - 1 To 0 Step -1
            If now_pid_list.IndexOf(ok_list(i)) = -1 Then
                ok_list.RemoveAt(i)
            End If
        Next
    End Sub

    Private Sub copy_timer_Tick(sender As Object, e As EventArgs)
        sender.stop()
        Dim p As Process = task_list(sender.tag)
        Dim cmdline = GetCommandLine(p)
        If System.Text.RegularExpressions.Regex.Matches(cmdline, """").Count Mod 2 = 1 Then
            cmdline += """"
        End If
        Dim res As String = ""
        Dim i, j As Int32
        i = 0
        j = -1
        Do
            i = cmdline.IndexOf("""", j + 1)
            If i = -1 Then Exit Do
            j = cmdline.IndexOf("""", i + 1)
            Dim t As String = cmdline.Substring(i + 1, j - i - 1)
            If IO.File.Exists(t) Then
                Dim tmp As New IO.FileInfo(t)
                If ext_list.IndexOf(tmp.Extension.ToLower) <> -1 Then
                    res = t
                    Exit Do
                End If
            End If
            If j = cmdline.Length - 1 Then Exit Do
        Loop

        If res <> "" Then
            Dim f As Boolean = True
            Dim tmp As New IO.FileInfo(res)
            If tmp.Length / 1024 > NumericUpDown2.Value Then
                f = False
            End If
            If GetDriveType(System.IO.Path.GetPathRoot(res)) <> DRIVE_REMOVABLE And CheckBox14.Checked = True Then
                f = False
            End If
            If f Then
                If copy_to_disk Then
                    For Each t As String In disk_save_paths
                        Dim new_name As String = tmp.Name.Replace(tmp.Extension, "") + Format(Now(), "_yyyy_MM_dd_HH_mm_ss_ff") + tmp.Extension
                        IO.File.Copy(res, t + new_name)
                    Next
                End If
            End If
        End If
        task_list.Remove(p)

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        load_listbox1()
        switch1()
    End Sub

    Private Sub switch1()
        If ti IsNot Nothing Then ti.Dispose()
        ti = New Timer
        ti.Interval = 20
        AddHandler ti.Tick, AddressOf cs
        slidetime = 0
        Panel1.Left = 0
        ti.Start()
    End Sub

    Private Sub cs(sender As Object, e As EventArgs)
        slidetime += 1
        Panel1.Left = smooth(slidetime / 25.0) * (-459)
        Panel2.Left = Panel1.Left + Panel1.Width
        If slidetime = 25 Then
            sender.stop()
            sender.dispose()
        End If
    End Sub

    Private Sub switch2()
        If ti IsNot Nothing Then ti.Dispose()
        ti = New Timer
        ti.Interval = 20
        AddHandler ti.Tick, AddressOf cs2
        slidetime = 0
        Panel2.Left = 0
        ti.Start()
    End Sub

    Private Sub cs2(sender As Object, e As EventArgs)
        slidetime += 1
        Panel1.Left = -459 + smooth(slidetime / 25.0) * 459
        Panel2.Left = Panel1.Left + Panel1.Width
        If slidetime = 25 Then
            sender.stop()
            sender.dispose()
        End If
    End Sub

    Private Sub load_listbox1()
        ListBox1.Items.Clear()
        For Each i In disk_save_paths
            ListBox1.Items.Add(i)
        Next

    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked Then
            copy_to_disk = True
        Else
            copy_to_disk = False
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Dim tmp As String = FolderBrowserDialog1.SelectedPath
            If tmp.Substring(tmp.Length - 1) <> "\" Then
                tmp = tmp & "\"
            End If
            disk_save_paths.Add(tmp)
            ListBox1.Items.Add(tmp)
            save_settings()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        switch2()
    End Sub

    Private Sub mini_timer_Tick(sender As Object, e As EventArgs)
        TryCast(sender, Timer).Stop()
        Me.WindowState = FormWindowState.Minimized
    End Sub
End Class
