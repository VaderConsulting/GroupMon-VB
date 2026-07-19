#Region " Imports "
Imports System.ServiceProcess
Imports System.IO
Imports System.Environment
Imports System.Configuration
Imports System.Windows.Forms
#End Region

Public Class svcMain

#Region " Inherits "

    Inherits System.ServiceProcess.ServiceBase

#End Region

#Region " Component Designer generated code "
    Friend WithEvents ApplicationEventLog As System.Diagnostics.EventLog

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents Control As System.ServiceProcess.ServiceController

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents tmrMain As System.Timers.Timer

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.tmrMain = New System.Timers.Timer
        Me.Control = New System.ServiceProcess.ServiceController
        Me.ApplicationEventLog = New System.Diagnostics.EventLog
        CType(Me.tmrMain, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ApplicationEventLog, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'tmrMain
        '
        Me.tmrMain.Interval = 60000
        '
        'Control
        '
        Me.Control.ServiceName = "Group Monitor Service"
        '
        'ApplicationEventLog
        '
        Me.ApplicationEventLog.Log = "Application"
        Me.ApplicationEventLog.Source = "Group Monitor Service"
        '
        'svcMain
        '
        Me.CanPauseAndContinue = True
        Me.ServiceName = "Group Monitor Service"
        CType(Me.tmrMain, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ApplicationEventLog, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New svcMain}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

#End Region

#Region " Public Variables "

    Public DebugMode As Boolean = False
    Public doPause As Boolean = False
    Public doStop As Boolean = False
    Public MonitoringInterval As Integer
    Public TabDepth As Integer
    Public DebugLog As StreamWriter
    Public Groups() As Group
    Public WithEvents Worker As New Worker
    Public AppName As String
    Public CurrentDir As String
    Public DebugLogfilePath As String
    Public DefaultMailTo As String
    Public Exclusions() As String
    Public MailFrom As String
    Public SMTPServer As String
    Public Subject As String

#End Region

#Region " Structures "

    Public Structure Group
        Dim Path As String
        Dim Members As System.Collections.SortedList
        Dim AlertEmailAddress As String
    End Structure

#End Region

#Region " Subroutines "

    Private Sub CreateSubdirectories()
        If Not My.Computer.FileSystem.DirectoryExists(CurrentDir & "Logs") Then
            My.Computer.FileSystem.CreateDirectory(CurrentDir & "Logs")
        End If

        If Not My.Computer.FileSystem.DirectoryExists(CurrentDir & "Results") Then
            My.Computer.FileSystem.CreateDirectory(CurrentDir & "Results")
        End If

        If Not My.Computer.FileSystem.DirectoryExists(CurrentDir & "Config") Then
            My.Computer.FileSystem.CreateDirectory(CurrentDir & "Config")
        End If
    End Sub

    Private Sub FinishLogging()
        DebugLog.Flush()
        DebugLog.Close()
    End Sub

    Private Sub GetConfig()
        Dim ConfigSettings As AppSettingsReader = New AppSettingsReader

        Try
            DebugMode = CType(ConfigSettings.GetValue("DebugMode", GetType(System.Boolean)), Boolean)
            AppName = CType(ConfigSettings.GetValue("AppName", GetType(System.String)), String)
            SMTPServer = CType(ConfigSettings.GetValue("smtpServer", GetType(System.String)), String)
            DefaultMailTo = CType(ConfigSettings.GetValue("DefaultMailTo", GetType(System.String)), String)
            MailFrom = CType(ConfigSettings.GetValue("MailFrom", GetType(System.String)), String)
            Subject = CType(ConfigSettings.GetValue("Subject", GetType(System.String)), String)
            MonitoringInterval = CType(ConfigSettings.GetValue("MonitoringIntervalMinutes", GetType(System.String)), Integer)
        Catch
            DebugMode = True
            AppName = "Group Monitoring Service"
            SMTPServer = "smtphost"
            DefaultMailTo = ""
            MailFrom = ""
            Subject = "Group Monitoring Alert"
            MonitoringInterval = 1
        End Try

        DebugLogfilePath = CurrentDir & "Logs\debug.log"

    End Sub

    Private Sub GetCurrentDir()
        ' Get the current directory, ie the directory where the application lives
        CurrentDir = Application.StartupPath

        If Not CurrentDir.EndsWith("\") Then
            CurrentDir += "\"
        End If
    End Sub

    Private Sub GetExclusions()
        Dim i As Integer

        WriteToDebugLog("Inside GetExclusions")

        If System.IO.File.Exists(CurrentDir & "Config" & "\Exclude.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(CurrentDir & "Config" & "\Exclude.txt")
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve Exclusions(i)
                    If strData Is Nothing Then Exit Do
                    If strData.Length > 0 And strData.Trim.StartsWith("#") = False Then
                        Exclusions(i) = strData

                        WriteToDebugLog("Loaded Exclusion:" & Exclusions(i))
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
        WriteToDebugLog("Leaving GetExclusions")
        WriteToDebugLog("--------------------------------------------------------------------------------")
    End Sub

    Private Sub GetGroupList()
        Dim i As Integer
        Dim strDataPortions() As String
        Dim strLeftPortion As String
        Dim strRightPortion As String

        WriteToDebugLog("Inside GetGroupList")

        If System.IO.File.Exists(CurrentDir & "Config\Grouplist.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(CurrentDir & "Config\Grouplist.txt")
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve Groups(i)
                    If strData Is Nothing Then Exit Do
                    If strData <> "" And strData.Trim.StartsWith("#") = False Then
                        strDataPortions = Split(strData, "|", -1, CompareMethod.Text)
                        strLeftPortion = strDataPortions(0)
                        Groups(i).Path = strLeftPortion
                        WriteToDebugLog("Loaded Group Path:" & Groups(i).Path)

                        ' Just in case there is no pipe (|), place the code to extract the 2nd portion inside try.catch.end try
                        Try
                            strRightPortion = strDataPortions(1)
                        Catch
                            strRightPortion = DefaultMailTo
                            WriteToDebugLog("No 'TO:' email address specified for group.  Using default.")
                        End Try

                        Groups(i).AlertEmailAddress = strRightPortion
                        WriteToDebugLog("Loaded Email Address:" & Groups(i).AlertEmailAddress)
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
        WriteToDebugLog("Leaving GetGroupList")
    End Sub

    Protected Overrides Sub OnContinue()

        WriteToDebugLog(AppName & " continued.")

        GetConfig()

        doPause = False

        Me.tmrMain.Enabled = True

    End Sub

    Protected Overrides Sub OnPause()

        WriteToDebugLog(AppName & " paused.")

        doPause = True

        Me.tmrMain.Enabled = False

    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        GetCurrentDir()

        GetConfig()

        CreateSubdirectories()

        doStop = False

        StartLogging()

        ' Output startup mode to eventlog
        WriteToDebugLog(AppName & " started.")

        ' Get the list of groups to monitor
        GetGroupList()

        ' Get the list of exclusions
        GetExclusions()

        WriteToDebugLog(UBound(Groups) & " Group(s) to monitor")

        Me.tmrMain.Enabled = True

    End Sub

    Protected Overrides Sub OnStop()
        WriteToDebugLog(AppName & " stopped.")

        doStop = True

        FinishLogging()

        Me.tmrMain.Enabled = False

    End Sub

    Private Sub oWorker_Alert(ByVal Message As String) Handles Worker.Alert
        WriteToDebugLog(Message)
    End Sub

    Private Sub StartLogging()
        If File.Exists(DebugLogfilePath) Then ' Rename the log
            ' Remove any old logs first
            If File.Exists(DebugLogfilePath & ".old") Then
                File.Delete(DebugLogfilePath & ".old")
            End If

            ' Rename (move) the file
            File.Move(DebugLogfilePath, DebugLogfilePath & ".old")
        End If
        DebugLog = File.CreateText(DebugLogfilePath)
    End Sub

    ' This is what actually kicks off the work...
    Private Sub tmrMain_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles tmrMain.Elapsed
        Dim oUser As DictionaryEntry
        Dim j As Int32
        Dim i As Int32
        Static iCounter As Integer

        iCounter += 1

        WriteToDebugLog("Timer fired.  Waiting for Counter to equal " & MonitoringInterval.ToString & ".  Counter now equals " & iCounter.ToString & ".")

        If iCounter < MonitoringInterval Then
            WriteToDebugLog("Leaving Timer subroutine.")
            Exit Sub
        End If

        iCounter = 0

        WriteToDebugLog("Inside Main Worker thread")
        WriteToDebugLog("--------------------------------------------------------------------------------")

        ' Disable timer to prevent recursive events from being triggered
        Me.tmrMain.Enabled = False

        ' For each group, get the list of users
        For i = 0 To UBound(Groups) - 1
            If Groups(i).Path <> "" Then
                WriteToDebugLog("Group Path: " & Groups(i).Path)

                ' Add exclusions
                For j = 0 To UBound(Exclusions)
                    If Exclusions(j) <> "" Then
                        Worker.Exclusions.Add(Exclusions(j), Exclusions(j))
                    End If
                Next j

                Worker.GroupPath = Groups(i).Path
                Worker.SMTPServer = SMTPServer
                Worker.MailFrom = MailFrom
                Worker.MailTo = Groups(i).AlertEmailAddress
                Worker.Subject = Subject
                Worker.OutputPath = CurrentDir & "Results\"

                WriteToDebugLog("Calling GetUsers")
                TabDepth += 4
                Worker.GetUsers()
                TabDepth -= 4

                Groups(i).Members = Worker.UserList

                WriteToDebugLog("Calling ClearLists")
                TabDepth += 4
                Worker.ClearLists()
                TabDepth -= 4
                WriteToDebugLog("--------------------------------------------------------------------------------")

                oUser = Nothing
            End If
        Next i

        ' Force Garbage Collection
        GC.Collect()
        GC.WaitForPendingFinalizers()

        WriteToDebugLog("Leaving Main Worker thread")
        WriteToDebugLog("################################################################################")

        ' Re-enable timer
        Me.tmrMain.Enabled = True
    End Sub

    Public Sub WriteToDebugLog(ByVal strMessage As String)
        If strMessage <> "" Then
            Try
                If DebugMode Then
                    DebugLog.WriteLine(Now & " " & Space(TabDepth) & strMessage)
                    DebugLog.Flush()
                End If
            Catch
                ' Doh.
            End Try
        End If
    End Sub

#End Region

End Class
