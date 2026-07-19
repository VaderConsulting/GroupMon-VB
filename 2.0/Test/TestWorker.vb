Imports System.Windows.Forms.Application

Public Class Worker

#Region " Events "

    Public Event Alert(ByVal Message As String, ByVal MsgType As MessageType)

#End Region

#Region " Private variables "

    Private moUserList As New System.Collections.SortedList
    Private moPreviousList As New System.Collections.SortedList
    Private moExclusionList As New System.Collections.SortedList
    Private mbConnectionError As Boolean = False
    Private mlErrorCode As Long
    Private mstrErrorDescription As String
    Private mstrGroupPath As String
    Private mstrFilename As String
    Private mstrOuputPath As String
    Private mstrSMTPServer As String
    Private mstrMailTo As String
    Private mstrMailFrom As String
    Private mstrSubject As String

#End Region

#Region " Properties "

    Public Property GroupPath() As String
        Get
            RaiseEvent Alert("Retrieving GroupPath", MessageType.Debug)
            Return mstrGroupPath
        End Get
        Set(ByVal Value As String)
            RaiseEvent Alert("Setting GroupPath", MessageType.Debug)
            mstrGroupPath = Value
        End Set
    End Property

    Public ReadOnly Property UserList() As System.Collections.SortedList
        Get
            Return moUserList
        End Get
    End Property

    Public ReadOnly Property bConnectionError() As Boolean
        Get
            Return mbConnectionError
        End Get
    End Property

    Public ReadOnly Property lErrorCode() As Long
        Get
            Return (mlErrorCode)
        End Get
    End Property

    Public ReadOnly Property strErrorDescription() As String
        Get
            Return mstrErrorDescription
        End Get
    End Property

    Public Property SMTPServer() As String
        Get
            Return mstrSMTPServer
        End Get
        Set(ByVal Value As String)
            mstrSMTPServer = Value
        End Set
    End Property

    Public Property MailTo() As String
        Get
            Return mstrMailTo
        End Get
        Set(ByVal Value As String)
            mstrMailTo = Value
        End Set
    End Property

    Public Property MailFrom() As String
        Get
            Return mstrMailFrom
        End Get
        Set(ByVal Value As String)
            mstrMailFrom = Value
        End Set
    End Property

    Public Property Subject() As String
        Get
            Return mstrSubject
        End Get
        Set(ByVal Value As String)
            mstrSubject = Value
        End Set
    End Property

    Public Property Exclusions() As System.Collections.SortedList
        Get
            Return moExclusionList
        End Get
        Set(ByVal Value As System.Collections.SortedList)
            moExclusionList = Value
        End Set
    End Property

    Public Property OutputPath() As String
        Get
            Return mstrOuputPath
        End Get
        Set(ByVal value As String)
            mstrOuputPath = value
        End Set
    End Property

#End Region

#Region " Subroutines "

    Public Sub ClearLists()
        moUserList.Clear()
        moPreviousList.Clear()
        moExclusionList.Clear()
    End Sub

    Private Sub GetAllMembers(ByVal strGroupPath As String)
        Dim oDS As New System.DirectoryServices.DirectorySearcher
        Dim oSRColl As System.DirectoryServices.SearchResultCollection
        Dim oSR As System.DirectoryServices.SearchResult
        Dim oDE As System.DirectoryServices.DirectoryEntry
        Dim i As Int32

        RaiseEvent Alert("Inside GetAllMembers", MessageType.Debug)

        Try
            oDS.SearchRoot.Path = strGroupPath
            oDS.PropertiesToLoad.Add("ADsPath")
            oDS.SearchScope = DirectoryServices.SearchScope.Subtree
            oDS.Filter = "(|(|(objectClass=group)(objectClass=user))(objectClass=computer))"

            oSRColl = oDS.FindAll()

            ' Clear existing members
            moUserList.Clear()

            For i = 0 To oSRColl.Count - 1
                oSR = oSRColl.Item(i)
                oDE = oSR.GetDirectoryEntry
                If oDE.SchemaClassName = "group" Then ' We are only monitoring groups
                    RaiseEvent Alert("Calling GetGroupMembers (" & oDE.Path & ")", MessageType.Debug)
                    GetGroupMembers(oDE.Path)
                End If
            Next

        Catch
            mbConnectionError = True
            mlErrorCode = Err.Number
            mstrErrorDescription = Err.Description
            Err.Clear()
            RaiseEvent Alert("Error " & mlErrorCode & " (" & mstrErrorDescription & ") in GetAllMembers.", MessageType.Normal)
        End Try

        oDS = Nothing
        oDE = Nothing
        oSR = Nothing
        oSRColl = Nothing

        RaiseEvent Alert("Leaving GetAllMembers", MessageType.Debug)

    End Sub

    Private Sub GetGroupMembers(ByVal strGroupPath As String)
        Dim oGroup As ActiveDs.IADsGroup
        Dim oUserOrGroup As ActiveDs.IADs
        Dim oMember As Object

        RaiseEvent Alert("Inside GetGroupMembers", MessageType.Debug)

        Try
            oGroup = GetObject(strGroupPath)
            DoEvents()

            For Each oUserOrGroup In oGroup.Members
                If oUserOrGroup.Class = "group" Then ' Call self
                    GetGroupMembers(oUserOrGroup.ADsPath)
                    For Each oMember In moUserList.Values
                        If Not moUserList.ContainsKey(oMember) Then
                            moUserList.Add(oMember, oMember)
                        End If
                    Next
                Else
                    If Not moUserList.ContainsKey(oUserOrGroup.ADsPath) Then
                        moUserList.Add(oUserOrGroup.ADsPath, oUserOrGroup.ADsPath)
                    End If
                End If
            Next

        Catch
            mbConnectionError = True
            mlErrorCode = Err.Number
            mstrErrorDescription = Err.Description
            Err.Clear()
            RaiseEvent Alert("Error " & mlErrorCode & " (" & mstrErrorDescription & ") in GetGroupMembers.", MessageType.Normal)
        End Try

        oGroup = Nothing
        oUserOrGroup = Nothing

        RaiseEvent Alert("Leaving GetGroupMembers", MessageType.Debug)
    End Sub

    Private Sub GetPreviousList()
        RaiseEvent Alert("Inside GetPreviousList", MessageType.Debug)

        If System.IO.File.Exists(mstrFilename) = True Then

            Dim oTextReader As New System.IO.StreamReader(mstrFilename)
            Dim oUsers As New Collections.SortedList
            Dim strData As String

            Try
                Do
                    strData = oTextReader.ReadLine()
                    If strData <> "" Then
                        oUsers.Add(strData, strData)
                    End If
                Loop Until strData Is Nothing

                oTextReader.Close()
            Catch
                mlErrorCode = Err.Number
                mstrErrorDescription = Err.Description
                Err.Clear()
                RaiseEvent Alert("Error " & mlErrorCode & " (" & mstrErrorDescription & ") in GetPreviousList.", MessageType.Normal)
            End Try
            moPreviousList = oUsers
        Else
            RaiseEvent Alert("No previous list found.  This must be a first run.", MessageType.Normal)
        End If

        RaiseEvent Alert("Leaving GetPreviousList", MessageType.Debug)
    End Sub

    Private Sub CompareLists()
        Dim oNewList As New System.Collections.SortedList
        Dim oOldList As New System.Collections.SortedList
        Dim oAddedUsers As New System.Collections.SortedList
        Dim oRemovedUsers As New System.Collections.SortedList
        Dim OldUser As String
        Dim NewUser As String
        Dim oUser As String
        Dim strMessage As String
        Dim bSendMail As Boolean
        Dim strExclusion As DictionaryEntry

        RaiseEvent Alert("Inside CompareLists", MessageType.Debug)

        oOldList = moPreviousList
        oNewList = moUserList
        OldUser = ""
        NewUser = ""
        oUser = ""

        For Each OldUser In oOldList.Values
            If oNewList.ContainsValue(OldUser) = False Then
                For Each strExclusion In moExclusionList
                    If OldUser.ToUpper Like strExclusion.Value.ToString.ToUpper = False Then
                        ' Kludge to remove SMS accounts
                        If OldUser.StartsWith(strExclusion.Value.ToString.ToUpper) = False Then
                            oRemovedUsers.Add(OldUser, OldUser)
                            RaiseEvent Alert("Added " & OldUser & " to Removed User List", MessageType.Debug)
                        End If
                        bSendMail = True
                    End If
                Next
            End If
        Next

        For Each NewUser In oNewList.Values
            If oOldList.ContainsValue(NewUser) = False Then
                For Each strExclusion In moExclusionList
                    If OldUser.ToUpper Like strExclusion.Value.ToString.ToUpper = False Then
                        ' Kludge to remove SMS accounts
                        If NewUser.StartsWith(strExclusion.Value.ToString.ToUpper) = False Then
                            oAddedUsers.Add(NewUser, NewUser)
                            RaiseEvent Alert("Added " & NewUser & " to Added User List", MessageType.Debug)
                        End If
                        bSendMail = True
                    End If
                Next
            End If
        Next

        strMessage = ""

        strMessage += "<HTML>"
        strMessage += "<BODY>"
        strMessage += "<p style='FONT-WEIGHT: bold; FONT-SIZE: 10pt; COLOR: black; FONT-FAMILY: Arial; TEXT-ALIGN: center'>Group Monitoring Report</p>"
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>This email has been sent to inform you of a change in the group membership of the following Group:</p>"
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'</p>" & mstrGroupPath & ""

        strMessage += "<hr>"

        If oAddedUsers.Count <> 0 Then
            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>The following users were added since the last runtime:</p>"
            RaiseEvent Alert("The following users were added since the last runtime:", MessageType.Normal)

            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
            For Each oUser In oAddedUsers.Keys
                strMessage += oUser & "<br>"
                RaiseEvent Alert(oUser, MessageType.Normal)
            Next
            strMessage += "</p>"
        Else
            RaiseEvent Alert("No users were added since the last runtime", MessageType.Normal)
        End If

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        If oRemovedUsers.Count <> 0 Then
            strMessage += "The following users were removed since the last runtime:</p>"
            RaiseEvent Alert("The following users were removed since the last runtime:", MessageType.Normal)

            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
            For Each oUser In oRemovedUsers.Keys
                strMessage += oUser & "<br>"
                RaiseEvent Alert(oUser, MessageType.Normal)
            Next
            strMessage += "</p>"
        Else
            RaiseEvent Alert("No users were removed since the last runtime", MessageType.Normal)
        End If

        strMessage += "<hr>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following users are now in this group:<br>"
        RaiseEvent Alert("The following users are now in this group:", MessageType.Normal)

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        For Each NewUser In oNewList.Values
            strMessage += NewUser & "<br>"
            RaiseEvent Alert(NewUser, MessageType.Normal)
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following exclusions are in force, and will not be reported:</p>"
        RaiseEvent Alert("The following exclusions are in force, and will not be reported:", MessageType.Normal)

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        For Each strExclusion In moExclusionList
            strMessage += strExclusion.Value.ToString & "<br>"
            RaiseEvent Alert(strExclusion.Value.ToString, MessageType.Normal)
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "Time of extraction: " & Now & "<br>"
        strMessage += "This report shows the EFFECTIVE permissions of a user.<br>"

        strMessage += "<br>The service that produced this report is running on " & Microsoft.VisualBasic.Environ("COMPUTERNAME") & "."

        strMessage += "</p>"

        strMessage += "</BODY>"
        strMessage += "</HTML>"

        'If bSendMail = True Then
        'RaiseEvent Alert("Calling SendEmail")
        'SendEmail(mstrMailFrom, mstrMailTo, mstrSubject & " - " & mstrGroupPath, strMessage, mstrSMTPServer, True)
        'End If

        RaiseEvent Alert("Leaving CompareLists", MessageType.Debug)

    End Sub

    Private Sub WriteOutput()
        Dim oUser As DictionaryEntry

        If System.IO.File.Exists(mstrFilename) Then
            If System.IO.File.Exists(mstrFilename & ".old") Then
                System.IO.File.Delete(mstrFilename & ".old")
            End If
            System.IO.File.Move(mstrFilename, mstrFilename & ".old")
        End If

        Dim oTextWriter As New System.IO.StreamWriter(mstrFilename)

        For Each oUser In moUserList
            oTextWriter.WriteLine(oUser.Key)
        Next

        oTextWriter.Flush()
        oTextWriter.Close()
    End Sub

    Private Function SendEmail(ByVal strFrom As String, ByVal strTo As String, ByVal strSubject As String, ByVal strMessage As String, ByVal strServer As String, Optional ByVal bHTML As Boolean = False) As Boolean
        Dim objEmail As Object

        RaiseEvent Alert("Inside SendEmail", MessageType.Debug)

        If Trim(strFrom) = "" Or Trim(strTo) = "" Or Trim(strSubject) = "" Or Trim(strMessage) = "" Or Trim(strServer) = "" Then
            Exit Function
        End If

        SendEmail = True

        objEmail = CreateObject("CDO.Message")

        objEmail.From = strFrom
        objEmail.To = strTo
        objEmail.Subject = strSubject

        If bHTML = True Then
            objEmail.Htmlbody = strMessage
        Else
            objEmail.Textbody = strMessage
        End If

        ' sendusing:
        ' 1 use local SMTP service (local to the specified smtpserver)
        ' 2 use port
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        objEmail.Configuration.Fields.Update()

        On Error Resume Next

        RaiseEvent Alert("Attempting to send email to " & strTo & ".", MessageType.Debug)
        objEmail.Send()
        objEmail = Nothing

        'Return True if it worked
        If Err.Number = 0 Then
            RaiseEvent Alert("Error sending email", MessageType.Debug)
            RaiseEvent Alert("Leaving SendEmail", MessageType.Debug)
            Return False
        Else
            RaiseEvent Alert("Email sent", MessageType.Debug)
            RaiseEvent Alert("Leaving SendEmail", MessageType.Debug)
            Return True
        End If
    End Function

    Public Sub GetUsers()
        RaiseEvent Alert("Inside GetUsers", MessageType.Debug)

        mbConnectionError = False
        mlErrorCode = 0&
        mstrErrorDescription = ""

        RaiseEvent Alert("Calling GetAllMembers", MessageType.Debug)
        GetAllMembers(mstrGroupPath)

        If mlErrorCode <> 0 Then
            RaiseEvent Alert("Error detected in GetAllMembers, exiting GetUsers", MessageType.Normal)
        Else
            RaiseEvent Alert("Calling GetFilename", MessageType.Debug)
            GetFilename()

            RaiseEvent Alert("Calling GetPreviousList", MessageType.Debug)
            GetPreviousList()

            RaiseEvent Alert("Calling CompareLists", MessageType.Debug)
            CompareLists()

            RaiseEvent Alert("Calling WriteOutput", MessageType.Debug)
            WriteOutput()

            moUserList = UserList

        End If

        RaiseEvent Alert("Leaving GetUsers", MessageType.Debug)

    End Sub

    Private Sub GetFilename()
        Dim strFilename As String
        Dim strPath As String
        Dim iLength As Int16

        RaiseEvent Alert("Inside GetFilename", MessageType.Debug)

        iLength = InStr(11, mstrGroupPath, "=", CompareMethod.Text)
        strFilename = mstrGroupPath.Substring(10, iLength - 14)

        strPath = mstrOuputPath

        If strPath.EndsWith("\") = False Then
            strPath += "\"
        End If

        strPath += strFilename & ".txt"

        mstrFilename = strPath

        RaiseEvent Alert("Leaving GetFilename", MessageType.Debug)
    End Sub

#End Region

#Region " Functions "


#End Region

End Class
