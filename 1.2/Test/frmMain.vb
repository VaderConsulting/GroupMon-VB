Imports System.ServiceProcess
Imports System.IO
Imports System.Data
Imports System.Environment
Imports System.Configuration
Imports System.Windows.Forms

Public Enum MessageType
    Debug = 0
    Normal = 1
End Enum

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Public Structure Group
        Dim Path As String
        Dim Members As System.Collections.SortedList
        Dim AlertEmailAddress As String
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
    Public strDefaultMailTo As String
    Public strMailFrom As String
    Public strSubject As String
    Public strExclusions() As String
    Friend WithEvents cmbGroupList As System.Windows.Forms.ComboBox
    Friend WithEvents chkShowDebugMessages As System.Windows.Forms.CheckBox
    Public intTabDepth As Integer

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Button1 = New System.Windows.Forms.Button
        Me.lstMain = New System.Windows.Forms.ListBox
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.cmbGroupList = New System.Windows.Forms.ComboBox
        Me.chkShowDebugMessages = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(516, 338)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 24)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Check"
        Me.Button1.Visible = False
        '
        'lstMain
        '
        Me.lstMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstMain.Location = New System.Drawing.Point(8, 40)
        Me.lstMain.Name = "lstMain"
        Me.lstMain.Size = New System.Drawing.Size(592, 251)
        Me.lstMain.TabIndex = 1
        '
        'tmrMain
        '
        Me.tmrMain.Interval = 30000
        '
        'cmbGroupList
        '
        Me.cmbGroupList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbGroupList.FormattingEnabled = True
        Me.cmbGroupList.Location = New System.Drawing.Point(8, 13)
        Me.cmbGroupList.Name = "cmbGroupList"
        Me.cmbGroupList.Size = New System.Drawing.Size(592, 21)
        Me.cmbGroupList.TabIndex = 2
        '
        'chkShowDebugMessages
        '
        Me.chkShowDebugMessages.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkShowDebugMessages.AutoSize = True
        Me.chkShowDebugMessages.Location = New System.Drawing.Point(8, 297)
        Me.chkShowDebugMessages.Name = "chkShowDebugMessages"
        Me.chkShowDebugMessages.Size = New System.Drawing.Size(139, 17)
        Me.chkShowDebugMessages.TabIndex = 3
        Me.chkShowDebugMessages.Text = "Show Debug Messages"
        Me.chkShowDebugMessages.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 317)
        Me.Controls.Add(Me.chkShowDebugMessages)
        Me.Controls.Add(Me.cmbGroupList)
        Me.Controls.Add(Me.lstMain)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Group Status"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DoWork()
    End Sub

    Private Sub DoWork()
        Dim i As Int32

        WriteToDebugLog("Inside Main Worker thread")
        WriteToDebugLog("--------------------------------------------------------------------------------")

        ' For each group, get the list of users
        For i = 0 To UBound(oGroups) - 1
            If oGroups(i).Path <> "" Then
                WriteToDebugLog("Group Path: " & oGroups(i).Path)

                MonitorGroup(i)

            End If
        Next i

        ' Force Garbage Collection
        GC.Collect()
        GC.WaitForPendingFinalizers()

        WriteToDebugLog("Leaving Main Worker thread")
        WriteToDebugLog("################################################################################")
    End Sub

    Private Sub MonitorGroup(ByVal intGroupID As Integer)
        oWorker.GroupPath = oGroups(intGroupID).Path
        oWorker.SMTPServer = strSMTPServer
        oWorker.MailFrom = strMailFrom
        oWorker.MailTo = oGroups(intGroupID).AlertEmailAddress
        oWorker.Subject = strSubject
        oWorker.OutputPath = strCurrentDir & "Results\"

        WriteToDebugLog("Calling GetUsers")
        intTabDepth += 4
        oWorker.GetUsers()
        intTabDepth -= 4

        oGroups(intGroupID).Members = oWorker.UserList

        WriteToDebugLog("Calling ClearLists")
        intTabDepth += 4
        oWorker.ClearLists()
        intTabDepth -= 4
        WriteToDebugLog("--------------------------------------------------------------------------------")

        'oUser = Nothing
    End Sub

    Public Sub WriteToDebugLog(ByVal strMessage As String, Optional ByVal msg As MessageType = MessageType.Debug)
        If strMessage <> "" Then
            Try
                If bDebugMode Or msg = MessageType.Normal Then
                    Me.lstMain.Items.Add(Space(intTabDepth) & strMessage)
                    Me.lstMain.Refresh()
                    oDebugLog.WriteLine(Now & " " & Space(intTabDepth) & strMessage)
                    oDebugLog.Flush()
                End If
            Catch
                ' Doh.
            End Try
        End If
    End Sub

    Private Sub oWorker_Alert(ByVal Message As String, Optional ByVal msgType As MessageType = MessageType.Debug) Handles oWorker.Alert
        WriteToDebugLog(Replace(Message, vbCrLf, ""), msgType)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Refresh()

        GetCurrentDir()

        GetConfig()

        CreateSubdirectories()

        StartLogging()

        GetGroupList()

        GetExclusions()

        CopyGroupListToCombobox()
    End Sub

    Private Sub GetConfig()

        strDebugLogfilePath = strCurrentDir & "Logs\debug.log"

    End Sub

    Private Sub GetCurrentDir()
        ' Get the current directory, ie the directory where the application lives
        strCurrentDir = Application.StartupPath

        If Not strCurrentDir.EndsWith("\") Then
            strCurrentDir += "\"
        End If
    End Sub

    Private Sub CreateSubdirectories()
        If Not My.Computer.FileSystem.DirectoryExists(strCurrentDir & "Logs") Then
            My.Computer.FileSystem.CreateDirectory(strCurrentDir & "Logs")
        End If

        If Not My.Computer.FileSystem.DirectoryExists(strCurrentDir & "Results") Then
            My.Computer.FileSystem.CreateDirectory(strCurrentDir & "Results")
        End If

        If Not My.Computer.FileSystem.DirectoryExists(strCurrentDir & "Config") Then
            My.Computer.FileSystem.CreateDirectory(strCurrentDir & "Config")
        End If
    End Sub

    Private Sub GetExclusions()
        Dim i As Integer

        WriteToDebugLog("Inside GetExclusions")

        If System.IO.File.Exists(strCurrentDir & "Config" & "\Exclude.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(strCurrentDir & "Config" & "\Exclude.txt")
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    ReDim Preserve strExclusions(i)
                    If strData Is Nothing Then Exit Do
                    If strData <> "" And Trim(strData.StartsWith("#")) = False Then
                        strExclusions(i) = strData

                        WriteToDebugLog("Loaded Exclusion:" & strExclusions(i))
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
        Dim strDataPortions() As String
        Dim strLeftPortion As String
        Dim strRightPortion As String

        WriteToDebugLog("Inside GetGroupList")

        If System.IO.File.Exists(strCurrentDir & "Config\Grouplist.txt") = True Then
            Dim oTextReader As New System.IO.StreamReader(strCurrentDir & "Config\Grouplist.txt")
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
                        WriteToDebugLog("Loaded Group Path:" & oGroups(i).Path)

                        ' Just in case there is no pipe (|), place the code to extract the 2nd portion inside try.catch.end try
                        Try
                            strRightPortion = strDataPortions(1)
                        Catch
                            strRightPortion = strDefaultMailTo
                            WriteToDebugLog("No 'TO:' email address specified for group.  Using default.")
                        End Try

                        oGroups(i).AlertEmailAddress = strRightPortion
                        WriteToDebugLog("Loaded Email Address:" & oGroups(i).AlertEmailAddress)
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

    Private Sub CopyGroupListToCombobox()
        Dim i As Integer

        For i = 0 To UBound(oGroups) - 1
            Me.cmbGroupList.Items.Add(oGroups(i).Path)
        Next
    End Sub

    Private Sub AddExclusions()
        Dim j As Int32

        For j = 0 To UBound(strExclusions)
            If strExclusions(j) <> "" Then
                oWorker.Exclusions.Add(strExclusions(j), strExclusions(j))
            End If
        Next j

    End Sub

    Private Sub cmbGroupList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGroupList.SelectedIndexChanged
        If Me.chkShowDebugMessages.Checked = True Then
            bDebugMode = True
        Else
            bDebugMode = False
        End If

        Me.lstMain.Items.Clear()

        AddExclusions()

        MonitorGroup(cmbGroupList.SelectedIndex)
    End Sub
End Class
