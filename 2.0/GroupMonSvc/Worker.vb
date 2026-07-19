Imports System.Windows.Forms.Application

Public Class Worker

#Region " Events "

    Public Event Alert(ByVal Message As String)

#End Region

#Region " Private variables "
    Private _ConnectionError As Boolean = False
    Private _ErrorCode As Long
    Private _ExclusionList As New System.Collections.SortedList
    Private _PreviousList As New System.Collections.SortedList

    Private _UserList As New System.Collections.SortedList
    Private _ErrorDescription As String
    Private _Filename As String
    Private _GroupPath As String
    Private _MailFrom As String
    Private _MailTo As String
    Private _OuputPath As String
    Private _SMTPServer As String
    Private _Subject As String

#End Region

#Region " Properties "

    Public ReadOnly Property ConnectionError() As Boolean
        Get
            Return _ConnectionError
        End Get
    End Property

    Public Property Exclusions() As System.Collections.SortedList
        Get
            Return _ExclusionList
        End Get
        Set(ByVal Value As System.Collections.SortedList)
            _ExclusionList = Value
        End Set
    End Property

    Public Property GroupPath() As String
        Get
            RaiseEvent Alert("Retrieving GroupPath")
            Return _GroupPath
        End Get
        Set(ByVal Value As String)
            RaiseEvent Alert("Setting GroupPath")
            _GroupPath = Value
        End Set
    End Property

    Public ReadOnly Property ErrorCode() As Long
        Get
            Return (_ErrorCode)
        End Get
    End Property

    Public Property MailFrom() As String
        Get
            Return _MailFrom
        End Get
        Set(ByVal Value As String)
            _MailFrom = Value
        End Set
    End Property

    Public Property MailTo() As String
        Get
            Return _MailTo
        End Get
        Set(ByVal Value As String)
            _MailTo = Value
        End Set
    End Property

    Public Property OutputPath() As String
        Get
            Return _OuputPath
        End Get
        Set(ByVal value As String)
            _OuputPath = value
        End Set
    End Property

    Public Property SMTPServer() As String
        Get
            Return _SMTPServer
        End Get
        Set(ByVal Value As String)
            _SMTPServer = Value
        End Set
    End Property

    Public ReadOnly Property ErrorDescription() As String
        Get
            Return _ErrorDescription
        End Get
    End Property

    Public Property Subject() As String
        Get
            Return _Subject
        End Get
        Set(ByVal Value As String)
            _Subject = Value
        End Set
    End Property

    Public ReadOnly Property UserList() As System.Collections.SortedList
        Get
            Return _UserList
        End Get
    End Property

#End Region

#Region " Subroutines "

    Public Sub ClearLists()
        _UserList.Clear()
        _PreviousList.Clear()
        _ExclusionList.Clear()
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

        RaiseEvent Alert("Inside CompareLists")

        oOldList = _PreviousList
        oNewList = _UserList
        OldUser = ""
        NewUser = ""
        oUser = ""

        For Each OldUser In oOldList.Values
            If oNewList.ContainsValue(OldUser) = False Then
                For Each strExclusion In _ExclusionList
                    If OldUser.ToUpper Like strExclusion.Value.ToString.ToUpper = False Then
                        If OldUser.StartsWith(strExclusion.Value.ToString) = False Then
                            oRemovedUsers.Add(OldUser, OldUser)
                        End If
                        bSendMail = True
                    End If
                Next
            End If
        Next

        For Each NewUser In oNewList.Values
            If oOldList.ContainsValue(NewUser) = False Then
                For Each strExclusion In _ExclusionList
                    If OldUser.ToUpper Like strExclusion.Value.ToString.ToUpper = False Then
                        If NewUser.StartsWith(strExclusion.Value.ToString) = False Then
                            oAddedUsers.Add(NewUser, NewUser)
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
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'</p>" & _GroupPath & ""

        strMessage += "<hr>"

        If oAddedUsers.Count <> 0 Then
            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>The following users were added since the last runtime:</p>"
            RaiseEvent Alert("The following users were added since the last runtime:")

            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
            For Each oUser In oAddedUsers.Keys
                strMessage += Replace(Replace(oUser, "CN=", ""), "\", "") & "<br>"
                RaiseEvent Alert(Replace(Replace(oUser, "CN=", ""), "\", ""))
            Next
            strMessage += "</p>"
        Else
            RaiseEvent Alert("No users were added since the last runtime")
        End If

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        If oRemovedUsers.Count <> 0 Then
            strMessage += "The following users were removed since the last runtime:</p>"
            RaiseEvent Alert("The following users were removed since the last runtime:")

            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
            For Each oUser In oRemovedUsers.Keys
                strMessage += Replace(Replace(oUser, "CN=", ""), "\", "") & "<br>"
                RaiseEvent Alert(Replace(Replace(oUser, "CN=", ""), "\", ""))
            Next
            strMessage += "</p>"
        Else
            RaiseEvent Alert("No users were removed since the last runtime")
        End If

        strMessage += "<hr>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following users are now in this group:<br>"
        RaiseEvent Alert("The following users are now in this group:")
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"

        For Each NewUser In oNewList.Values
            strMessage += Replace(Replace(NewUser, "CN=", ""), "\", "") & "<br>"
            RaiseEvent Alert(Replace(Replace(NewUser, "CN=", ""), "\", ""))
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following exclusions are in force, and were not reported:</p>"
        RaiseEvent Alert("The following exclusions are in force, and were not reported:")

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        For Each strExclusion In _ExclusionList
            strMessage += Replace(strExclusion.Value.ToString, "LDAP://CN=", "") & "<br>"
            RaiseEvent Alert(Replace(strExclusion.Value.ToString, "LDAP://CN=", ""))
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "Time of extraction: " & Now & "<br>"
        strMessage += "This report shows Users only, but in a manner such that ANY User with these rights, either directly or indirectly is listed.<br>"

        strMessage += "<br>The service that produced this report is running on " & Microsoft.VisualBasic.Environ("COMPUTERNAME") & "."

        strMessage += "</p>"

        strMessage += "</BODY>"
        strMessage += "</HTML>"

        If bSendMail = True Then
            RaiseEvent Alert("Calling SendEmail")
            SendEmail(_MailFrom, _MailTo, _Subject & " - " & _GroupPath, strMessage, _SMTPServer, True)
        End If

        RaiseEvent Alert("Leaving CompareLists")

    End Sub

    Private Sub GetAllMembers(ByVal strGroupPath As String)
        Dim oDS As New System.DirectoryServices.DirectorySearcher
        Dim oSRColl As System.DirectoryServices.SearchResultCollection
        Dim oSR As System.DirectoryServices.SearchResult
        Dim oDE As System.DirectoryServices.DirectoryEntry
        Dim i As Int32

        RaiseEvent Alert("Inside GetAllMembers")

        Try
            oDS.SearchRoot.Path = strGroupPath
            oDS.PropertiesToLoad.Add("ADsPath")
            oDS.SearchScope = DirectoryServices.SearchScope.Subtree
            oDS.Filter = "(|(|(objectClass=group)(objectClass=user))(objectClass=computer))"

            oSRColl = oDS.FindAll()

            ' Clear existing members
            _UserList.Clear()

            For i = 0 To oSRColl.Count - 1
                oSR = oSRColl.Item(i)
                oDE = oSR.GetDirectoryEntry
                If oDE.SchemaClassName = "group" Then ' We are only monitoring groups
                    RaiseEvent Alert("Calling GetGroupMembers (" & oDE.Path & ")")
                    GetGroupMembers(oDE.Path)
                End If
            Next

        Catch
            _ConnectionError = True
            _ErrorCode = Err.Number
            _ErrorDescription = Err.Description
            Err.Clear()
            RaiseEvent Alert("Error " & _ErrorCode & " (" & _ErrorDescription & ") in GetAllMembers.")
        End Try

        oDS = Nothing
        oDE = Nothing
        oSR = Nothing
        oSRColl = Nothing

        RaiseEvent Alert("Leaving GetAllMembers")

    End Sub

    Private Sub GetFilename()
        Dim strFilename As String
        Dim strPath As String
        Dim iLength As Int32

        RaiseEvent Alert("Inside GetFilename")

        iLength = InStr(11, _GroupPath, "=", CompareMethod.Text)
        strFilename = _GroupPath.Substring(10, iLength - 14)

        strPath = _OuputPath

        If strPath.EndsWith("\") = False Then
            strPath += "\"
        End If

        strPath += strFilename & ".txt"

        _Filename = strPath

        RaiseEvent Alert("Leaving GetFilename")
    End Sub

    Private Sub GetGroupMembers(ByVal strGroupPath As String)
        Dim oGroup As ActiveDs.IADsGroup
        Dim oUserOrGroup As ActiveDs.IADs
        Dim oMember As Object

        RaiseEvent Alert("Inside GetGroupMembers")

        Try
            oGroup = GetObject(strGroupPath)
            DoEvents()

            For Each oUserOrGroup In oGroup.Members
                If oUserOrGroup.Class = "group" Then ' Call self
                    GetGroupMembers(oUserOrGroup.ADsPath)
                    For Each oMember In _UserList.Values
                        If Not _UserList.ContainsKey(oMember) Then
                            _UserList.Add(oMember, oMember)
                        End If
                    Next
                Else
                    If Not _UserList.ContainsKey(oUserOrGroup.ADsPath) Then
                        _UserList.Add(oUserOrGroup.Name, oUserOrGroup.Name)
                    End If
                End If
            Next

        Catch
            _ConnectionError = True
            _ErrorCode = Err.Number
            _ErrorDescription = Err.Description
            Err.Clear()
            RaiseEvent Alert("Error " & _ErrorCode & " (" & _ErrorDescription & ") in GetGroupMembers.")
        End Try

        oGroup = Nothing
        oUserOrGroup = Nothing

        RaiseEvent Alert("Leaving GetGroupMembers")
    End Sub

    Private Sub GetPreviousList()
        RaiseEvent Alert("Inside GetPreviousList")

        If System.IO.File.Exists(_Filename) = True Then

            Dim oTextReader As New System.IO.StreamReader(_Filename)
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
                _ErrorCode = Err.Number
                _ErrorDescription = Err.Description
                Err.Clear()
                RaiseEvent Alert("Error " & _ErrorCode & " (" & _ErrorDescription & ") in GetPreviousList.")
            End Try
            _PreviousList = oUsers
        Else
            RaiseEvent Alert("No previous list found.  This must be a first run.")
        End If

        RaiseEvent Alert("Leaving GetPreviousList")
    End Sub

    Public Sub GetUsers()
        RaiseEvent Alert("Inside GetUsers")

        _ConnectionError = False
        _ErrorCode = 0&
        _ErrorDescription = ""

        RaiseEvent Alert("Calling GetAllMembers")
        GetAllMembers(_GroupPath)

        If _ErrorCode <> 0 Then
            RaiseEvent Alert("Error detected in GetAllMembers, exiting GetUsers")
        Else
            RaiseEvent Alert("Calling GetFilename")
            GetFilename()

            RaiseEvent Alert("Calling GetPreviousList")
            GetPreviousList()

            RaiseEvent Alert("Calling CompareLists")
            CompareLists()

            RaiseEvent Alert("Calling WriteOutput")
            WriteOutput()

            _UserList = UserList

        End If

        RaiseEvent Alert("Leaving GetUsers")

    End Sub

    Private Sub WriteOutput()
        Dim oUser As DictionaryEntry

        If System.IO.File.Exists(_Filename) Then
            If System.IO.File.Exists(_Filename & ".old") Then
                System.IO.File.Delete(_Filename & ".old")
            End If
            System.IO.File.Move(_Filename, _Filename & ".old")
        End If

        Dim oTextWriter As New System.IO.StreamWriter(_Filename)

        For Each oUser In _UserList
            oTextWriter.WriteLine(oUser.Key)
        Next

        oTextWriter.Flush()
        oTextWriter.Close()
    End Sub

#End Region

#Region " Functions "

    Private Function SendEmail(ByVal strFrom As String, ByVal strTo As String, ByVal strSubject As String, ByVal strMessage As String, ByVal strServer As String, Optional ByVal bHTML As Boolean = False) As Boolean
        Dim Email As Object

        RaiseEvent Alert("Inside SendEmail")

        If Trim(strFrom) = "" Or Trim(strTo) = "" Or Trim(strSubject) = "" Or Trim(strMessage) = "" Or Trim(strServer) = "" Then
            Exit Function
        End If

        SendEmail = True

        Email = CreateObject("CDO.Message")

        Email.From = strFrom
        Email.To = strTo
        Email.Subject = strSubject

        If bHTML = True Then
            Email.Htmlbody = strMessage
        Else
            Email.Textbody = strMessage
        End If

        ' sendusing:
        ' 1 use local SMTP service (local to the specified smtpserver)
        ' 2 use port
        Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
        Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        Email.Configuration.Fields.Update()

        On Error Resume Next

        RaiseEvent Alert("Attempting to send email to " & strTo & ".")
        Email.Send()
        Email = Nothing

        'Return True if it worked
        If Err.Number = 0 Then
            RaiseEvent Alert("Email sent")
            RaiseEvent Alert("Leaving SendEmail")
            Return True
        Else
            RaiseEvent Alert("Error sending email")
            RaiseEvent Alert("Leaving SendEmail")
            Return False
        End If
    End Function

#End Region

End Class
