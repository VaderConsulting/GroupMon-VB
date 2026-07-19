Imports System.ServiceProcess
Imports System.IO
Imports System.Data
Imports System.Environment
Imports System.Configuration
Imports System.Windows.Forms

Public Class Form1
    Inherits System.Windows.Forms.Form

    Public Structure Group
        Dim Path As String
        Dim Members As System.Collections.SortedList
    End Structure

    Public bDebugMode As Boolean = False
    Public doPause As Boolean = False
    Public doStop As Boolean = False
    Public strCurrentDir As String
    Public strDebugLogfilePath As String
    Public strAppName As String
    Public oDebugLog As StreamWriter
    Public oGroups() As Group
    Public WithEvents oWorker As New Worker
    Public strSMTPServer As String
    Public strMailTo As String
    Public strMailFrom As String
    Public strSubject As String
    Public strExclusions() As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lstMain As System.Windows.Forms.ListBox
    Friend WithEvents tmrMain As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button1 = New System.Windows.Forms.Button
        Me.lstMain = New System.Windows.Forms.ListBox
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 24)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Start"
        '
        'lstMain
        '
        Me.lstMain.Location = New System.Drawing.Point(8, 40)
        Me.lstMain.Name = "lstMain"
        Me.lstMain.Size = New System.Drawing.Size(592, 303)
        Me.lstMain.TabIndex = 1
        '
        'tmrMain
        '
        Me.tmrMain.Enabled = True
        Me.tmrMain.Interval = 30000
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 374)
        Me.Controls.Add(Me.lstMain)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DoWork()
    End Sub

    Private Sub DoWork()
        Dim oGroup As DictionaryEntry
        Dim oUser As DictionaryEntry
        Dim strLogOutput As String
        Dim j As Int32
        Dim i As Int32

        WriteToDebugLog("[tmrMain_Elapsed]")

        ' Disable timer to prevent recursive events from being triggered
        Me.tmrMain.Enabled = False

        ' For each group, get the list of users
        For i = 0 To UBound(oGroups)
            If oGroups(i).Path <> "" Then
                WriteToDebugLog(oGroups(i).Path)

                ' Add exclusions
                For j = 0 To UBound(strExclusions)
                    If strExclusions(j) <> "" Then
                        oWorker.Exclusions.Add(strExclusions(j), strExclusions(j))
                    End If
                Next j

                oWorker.GroupPath = oGroups(i).Path
                oWorker.SMTPServer = strSMTPServer
                oWorker.MailFrom = strMailFrom
                oWorker.MailTo = strMailTo
                oWorker.Subject = strSubject

                WriteToDebugLog("Calling oWorker.GetUsers()")
                oWorker.GetUsers()

                WriteToDebugLog("Referencing oWorker.UserList")
                oGroups(i).Members = oWorker.UserList

                WriteToDebugLog("Calling oWorker.ClearLists()")
                oWorker.ClearLists()
                WriteToDebugLog("--------------------------------------------------------------------------------")

                oUser = Nothing
            End If
        Next i

        ' Force Garbage Collection
        GC.Collect()
        GC.WaitForPendingFinalizers()

        WriteToDebugLog("Run complete")
        WriteToDebugLog("################################################################################")

        ' Re-enable timer
        Me.tmrMain.Enabled = True
    End Sub

    Public Sub WriteToDebugLog(ByVal strMessage As String)
        'Dim f As FileInfo
        'Dim DebugFile As File

        'DebugFile.OpenRead(strDebugLogfilePath)

        If bDebugMode Then
            oDebugLog.WriteLine(Now & " " & strMessage)
            oDebugLog.Flush()
        End If

    End Sub

    Private Sub oWorker_Alert(ByVal Message As String) Handles oWorker.Alert
        WriteToDebugLog(Message)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetConfig()

        strDebugLogfilePath = strCurrentDir & "debug.log"

        StartLogging()

        GetGroupList()

        GetExclusions()

    End Sub

    Private Sub GetConfig()
        Dim ConfigSettings As AppSettingsReader = New AppSettingsReader

        Try
            bDebugMode = CType(ConfigSettings.GetValue("DebugMode", GetType(System.Boolean)), Boolean)
            strAppName = CType(ConfigSettings.GetValue("AppName", GetType(System.String)), String)
            strSMTPServer = CType(ConfigSettings.GetValue("smtpServer", GetType(System.String)), String)
            strMailTo = CType(ConfigSettings.GetValue("MailTo", GetType(System.String)), String)
            strMailFrom = CType(ConfigSettings.GetValue("MailFrom", GetType(System.String)), String)
            strSubject = CType(ConfigSettings.GetValue("Subject", GetType(System.String)), String)
        Catch
            'ApplicationEventLog.WriteEntry("Error " & Err.Number & " (" & Err.Description & ") in GetConfig")
            bDebugMode = True
            strAppName = "Group Monitor Service"
            strSMTPServer = "smtphost"
            strMailTo = ""
            strMailFrom = ""
            strSubject = "Group Monitoring Alert"
        End Try

        GetCurrentDir()

        strDebugLogfilePath = strCurrentDir & "debug.log"

    End Sub

    Private Sub GetCurrentDir()
        ' Get the current directory, ie the directory where the application lives
        strCurrentDir = Application.StartupPath

        If Not strCurrentDir.EndsWith("\") Then
            strCurrentDir += "\"
        End If
    End Sub

    Private Sub GetExclusions()
        Dim i As Integer

        WriteToDebugLog("[GetExclusions]")

        If System.IO.File.Exists(strCurrentDir & "Exclude.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(strCurrentDir & "Exclude.txt")
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve strExclusions(i)
                    If strData Is Nothing Then Exit Do
                    If strData <> "" And Trim(strData.StartsWith("#")) = False Then
                        strExclusions(i) = strData
                        i += 1
                    End If
                Loop Until strData Is Nothing
            Catch
                WriteToDebugLog("Error " & Err.Number & " (" & Err.Description & ") in GetExclusions.")
            End Try
            oTextReader.Close()
        Else
            WriteToDebugLog("No exclusions specified.")
        End If
    End Sub

    Private Sub StartLogging()
        If File.Exists(strDebugLogfilePath) Then ' Rename the log
            ' Remove any old logs first
            If File.Exists(strDebugLogfilePath & ".old") Then
                File.Delete(strDebugLogfilePath & ".old")
            End If

            ' Rename (move) the file
            File.Move(strDebugLogfilePath, strDebugLogfilePath & ".old")
        End If
        oDebugLog = File.CreateText(strDebugLogfilePath)
    End Sub

    Private Sub GetGroupList()
        Dim i As Integer

        WriteToDebugLog("[GetGroupList]")

        If System.IO.File.Exists(strCurrentDir & "Grouplist.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(strCurrentDir & "Grouplist.txt")
            Dim oUsers As New Collections.SortedList
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve oGroups(i)
                    If strData Is Nothing Then Exit Do
                    If strData <> "" And Trim(strData.StartsWith("#")) = False Then
                        oGroups(i).Path = strData
                        i += 1
                    End If
                Loop Until strData Is Nothing
            Catch
                WriteToDebugLog("Error " & Err.Number & " (" & Err.Description & ") in GetGroupList.")
            End Try
            oTextReader.Close()
        Else
            WriteToDebugLog("No groups specified.")
        End If
    End Sub

    Private Sub tmrMain_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMain.Tick
        DoWork()
    End Sub
End Class
