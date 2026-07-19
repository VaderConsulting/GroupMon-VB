Imports System.Windows.Forms.Application

Public Class Worker

#Region " Events "

    Public Event Alert(ByVal Message As String)

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
    Private mstrSMTPServer As String
    Private mstrMailTo As String
    Private mstrMailFrom As String
    Private mstrSubject As String

#End Region

#Region " Properties "

    Public Property GroupPath() As String
        Get
            RaiseEvent Alert("Retrieving GroupPath")
            Return mstrGroupPath
        End Get
        Set(ByVal Value As String)
            RaiseEvent Alert("Setting GroupPath")
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

        On Error Resume Next

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
                RaiseEvent Alert("Entering GetGroupMembers (" & oDE.Path & ")")
                GetGroupMembers(oDE.Path)
            End If
        Next

        If Err.Number <> 0 Then
            mbConnectionError = True
            mlErrorCode = Err.Number
            mstrErrorDescription = Err.Description
            Err.Clear()
            RaiseEvent Alert("Error " & Err.Number & " (" & Err.Description & ") in GetAllMembers.")
        End If

        oDS = Nothing
        oDE = Nothing
        oSR = Nothing
        oSRColl = Nothing

    End Sub

    Private Sub GetGroupMembers(ByVal strGroupPath As String)
        Dim oGroup As ActiveDs.IADsGroup
        Dim oUserOrGroup As ActiveDs.IADs
        Dim oMember As Object

        On Error Resume Next

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

        If Err.Number <> 0 Then
            mbConnectionError = True
            mlErrorCode = Err.Number
            mstrErrorDescription = Err.Description
            Err.Clear()
        End If

        oGroup = Nothing
        oUserOrGroup = Nothing
    End Sub

    Private Sub GetPreviousList()
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
                RaiseEvent Alert("Error " & Err.Number & " (" & Err.Description & ") in GetPreviousList.")
            End Try
            moPreviousList = oUsers
        End If
    End Sub

    Private Sub CompareLists()
        Dim oNewList As New System.Collections.SortedList
        Dim oOldList As New System.Collections.SortedList
        Dim oAddedUsers As New System.Collections.SortedList
        Dim oRemovedUsers As New System.Collections.SortedList
        Dim OldUser As String
        Dim NewUser As String
        Dim oUser As String
        Dim bFoundOldUserInNewList As Boolean
        Dim bFoundNewUserInOldList As Boolean
        Dim strMessage As String
        Dim bSendMail As Boolean
        Dim strExclusion As DictionaryEntry

        oOldList = moPreviousList
        oNewList = moUserList

        For Each OldUser In oOldList.Values
            If oNewList.ContainsValue(OldUser) = False Then
                For Each strExclusion In moExclusionList
                    If OldUser.ToUpper Like strExclusion.Value.ToString.ToUpper = False Then
                        ' Kludge to remove SMS accounts
                        If OldUser.StartsWith("LDAP://CN=SMS") = False Then
                            oRemovedUsers.Add(OldUser, OldUser)
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
                        If NewUser.StartsWith("LDAP://CN=SMS") = False Then
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
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'</p>" & mstrGroupPath & ""

        strMessage += "<hr>"

        If oAddedUsers.Count <> 0 Then
            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>The following users were added since the last runtime:</p>"
            RaiseEvent Alert("The following users were added since the last runtime:")

            strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
            For Each oUser In oAddedUsers.Keys
                strMessage += oUser & "<br>"
                RaiseEvent Alert(oUser)
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
                strMessage += oUser & "<br>"
                RaiseEvent Alert(oUser)
            Next
            strMessage += "</p>"
        Else
            RaiseEvent Alert("No users were removed since the last runtime")
        End If

        strMessage += "<hr>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following users are now in this group:<br>"
        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        For Each NewUser In oNewList.Values
            strMessage += NewUser & "<br>"
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "The following exclusions are in force, and will not be reported:</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 7pt; COLOR: red; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        For Each strExclusion In moExclusionList
            strMessage += strExclusion.Value.ToString & "<br>"
            RaiseEvent Alert(oUser)
        Next

        strMessage += "</p>"

        strMessage += "<p style='FONT-WEIGHT: normal; FONT-SIZE: 8pt; COLOR: blue; FONT-FAMILY: Arial; TEXT-ALIGN: left'>"
        strMessage += "Time of extraction: " & Now & "<br>"
        strMessage += "This report shows the EFFECTIVE permissions of a user.<br>"

        strMessage += "<br>The service that produced this report is running on " & Microsoft.VisualBasic.Environ("COMPUTERNAME") & "."

        strMessage += "</p>"

        strMessage += "</BODY>"
        strMessage += "</HTML>"

        If bSendMail = True Then
            SendEmail(mstrMailFrom, mstrMailTo, mstrSubject & " - " & mstrGroupPath, strMessage, mstrSMTPServer, True)
        End If

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
        objEmail.Send()

        'Return True if it worked
        If Err.Number = 0 Then
            Return False
        Else
            Return True
        End If
        On Error GoTo 0

        objEmail = Nothing
    End Function

    Public Sub GetUsers()
        RaiseEvent Alert("[GetUsers]")

        mbConnectionError = False
        mlErrorCode = 0&
        mstrErrorDescription = ""

        RaiseEvent Alert("Entering GetAllMembers")
        GetAllMembers(mstrGroupPath)

        If mlErrorCode <> 0 Then
            RaiseEvent Alert("Error, aborting")
        Else

            RaiseEvent Alert("Entering GetFilename")
            GetFilename()

            RaiseEvent Alert("Entering GetPreviousList")
            GetPreviousList()

            RaiseEvent Alert("Entering CompareLists")
            CompareLists()

            RaiseEvent Alert("Entering WriteOutput")
            WriteOutput()

            moUserList = UserList

        End If

    End Sub

#End Region

#Region " Functions "

    Private Function GetFilename() As String
        Dim strFilename As String
        Dim strPath As String
        Dim iLength As Int16

        iLength = InStr(11, mstrGroupPath, "=", CompareMethod.Text)
        strFilename = mstrGroupPath.Substring(10, iLength - 14)

        strPath = System.Windows.Forms.Application.StartupPath

        If strPath.EndsWith("\") = False Then
            strPath += "\"
        End If

        strPath += strFilename & ".txt"

        mstrFilename = strPath
    End Function

#End Region

End Class
