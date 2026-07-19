#Region " Imports "
Imports System.ServiceProcess
Imports System.IO
Imports System.Data
Imports System.Environment
Imports System.Configuration
Imports System.Windows.Forms
#End Region

Public Class svcMain

#Region " Inherits "

    Inherits System.ServiceProcess.ServiceBase

#End Region

#Region " Structures "

    Public Structure Group
        Dim Path As String
        Dim Members As System.Collections.SortedList
        Dim AlertEmailAddress As String
    End Structure

#End Region

#Region " Public Variables "

    Public bDebugMode As Boolean = False
    Public doPause As Boolean = False
    Public doStop As Boolean = False
    Public strCurrentDir As String
    Public strDebugLogfilePath As String
    Public strAppName As String
    Public oDebugLog As StreamWriter
    Public oGroups() As Group
    Public strSMTPServer As String
    Public strDefaultMailTo As String
    Public strMailFrom As String
    Public strSubject As String
    Public WithEvents oWorker As New Worker
    Public strExclusions() As String

#End Region

#Region " Component Designer generated code "

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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents tmrMain As System.Timers.Timer
    Friend WithEvents Control As System.ServiceProcess.ServiceController
    Friend WithEvents ApplicationEventLog As System.Diagnostics.EventLog

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
        Me.tmrMain.Enabled = False
        Me.tmrMain.Interval = CType(configurationAppSettings.GetValue("tmrMain.Interval", GetType(System.Double)), Double)
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

#End Region

#Region " Subroutines "

    Protected Overrides Sub OnStart(ByVal args() As String)
        GetConfig()

        GetCurrentDir()

        doStop = False

        StartLogging()

        ' Output startup mode to eventlog
        WriteToDebugLog(strAppName & " started.")

        ' Blurb
        WriteToDebugLog("Users in the following groups have been determined by collecting the users in each group,")
        WriteToDebugLog("as well as the users that are in each group WITHIN that group (if any), and so on until no more groups exist.")
        WriteToDebugLog(vbCrLf)
        WriteToDebugLog("What is thus presented is the Effective group membership.")
        WriteToDebugLog(vbCrLf)

        ' Get the list of groups to monitor
        GetGroupList()

        ' Get the list of exclusions
        GetExclusions()

        WriteToDebugLog(UBound(oGroups) & " Groups to monitor")

        Me.tmrMain.Enabled = True

    End Sub

    Protected Overrides Sub OnStop()
        WriteToDebugLog("Group Monitor Service stopped")

        doStop = True

        FinishLogging()

        Me.tmrMain.Enabled = False

    End Sub

    Protected Overrides Sub OnPause()

        WriteToDebugLog("Group Monitor Service paused")

        doPause = True

        Me.tmrMain.Enabled = False

    End Sub

    Protected Overrides Sub OnContinue()

        WriteToDebugLog("Group Monitor Service continued")

        GetConfig()

        doPause = False

        Me.tmrMain.Enabled = True

    End Sub

    Private Sub GetConfig()
        Dim ConfigSettings As AppSettingsReader = New AppSettingsReader

        Try
            bDebugMode = CType(ConfigSettings.GetValue("DebugMode", GetType(System.Boolean)), Boolean)
            strAppName = CType(ConfigSettings.GetValue("AppName", GetType(System.String)), String)
            strSMTPServer = CType(ConfigSettings.GetValue("smtpServer", GetType(System.String)), String)
            strDefaultMailTo = CType(ConfigSettings.GetValue("DefaultMailTo", GetType(System.String)), String)
            strMailFrom = CType(ConfigSettings.GetValue("MailFrom", GetType(System.String)), String)
            strSubject = CType(ConfigSettings.GetValue("Subject", GetType(System.String)), String)
        Catch
            ApplicationEventLog.WriteEntry("Error " & Err.Number & " (" & Err.Description & ") in GetConfig")
            bDebugMode = True
            strAppName = "Group Monitoring Service"
            strSMTPServer = "smtphost"
            strDefaultMailTo = ""
            strMailFrom = ""
            strSubject = "Group Monitoring Alert"
        End Try

        GetCurrentDir()

        GetExclusions()

        Me.ServiceName = strAppName

        strDebugLogfilePath = strCurrentDir & "debug.log"

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

    Private Sub FinishLogging()
        oDebugLog.Flush()
        oDebugLog.Close()
    End Sub

    Private Sub GetGroupList()
        Dim i As Integer
        Dim strDataPortions() As String
        Dim strLeftPortion As String
        Dim strRightPortion As String

        WriteToDebugLog("[GetGroupList]")

        If System.IO.File.Exists(strCurrentDir & "Grouplist.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(strCurrentDir & "Grouplist.txt")
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve oGroups(i)
                    If strData Is Nothing Then Exit Do
                    If strData <> "" And Trim(strData.StartsWith("#")) = False Then
                        strDataPortions = Split(strData, "|", -1, CompareMethod.Text)
                        strLeftPortion = strDataPortions(0)
                        oGroups(i).Path = strLeftPortion

                        ' Just in case there is no pipe (|), place the code to extract the 2nd portion inside try.catch.end try
                        Try
                            strRightPortion = strDataPortions(1)
                        Catch
                            strRightPortion = strDefaultMailTo
                            WriteToDebugLog("No 'TO:' email address specified for group.  Using default.")
                        End Try

                        oGroups(i).AlertEmailAddress = strRightPortion
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
        Try
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
                    WriteToDebugLog("Error " & Err.Number & " (" & Err.Description & ") in GetExclusions (Inner).")
                End Try
                oTextReader.Close()
            Else
                WriteToDebugLog("No exclusions specified.")
            End If
        Catch
            WriteToDebugLog("Error " & Err.Number & " (" & Err.Description & ") in GetExclusions (Outer).")
        End Try
        Err.Clear()
    End Sub

    ' This is what actually kicks off the work...
    Private Sub tmrMain_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles tmrMain.Elapsed
        Dim oUser As DictionaryEntry
        Dim i As Int32
        Dim j As Int32

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
                oWorker.MailTo = oGroups(i).AlertEmailAddress
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
        If strMessage <> "" Then
            Try
                If bDebugMode Then
                    oDebugLog.WriteLine(Now & " " & strMessage)
                    oDebugLog.Flush()
                End If
            Catch
                ' Doh.
            End Try
        End If
    End Sub

    Private Sub oWorker_Alert(ByVal Message As String) Handles oWorker.Alert
        WriteToDebugLog(Message)
    End Sub

#End Region

End Class
