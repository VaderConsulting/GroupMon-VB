Public Class frmStatus

    Dim oUserList As Object

    Private Sub ButtonQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonQuery.Click
        Dim Groupname As String
        Dim User As Object

        Me.Cursor = Cursors.WaitCursor

        Me.lstResults.Items.Clear()

        oUserList = CreateObject("Scripting.Dictionary")
        Groupname = My.Settings.GroupName

        GetUsers(Groupname)

        For Each User In oUserList
            Me.lstResults.Items.Add(Replace(User, "\", ""))
        Next

        Me.lblStatus.Text = "Last query: " & Now

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub GetUsers(ByVal Groupname)
        Dim oGroup As Object
        Dim Member As Object

        oGroup = GetObject(Groupname)

        For Each Member In oGroup.members
            If Microsoft.VisualBasic.Left(Member.ObjectCategory, 8) = "CN=Group" Then
                GetUsers(Member.aDSPath)
            Else
                On Error Resume Next ' Do this so that when we try and add the same user twice it doesn't balk
                oUserList.Add(Microsoft.VisualBasic.Right(Member.Name, Len(Member.Name) - 3), Member.Name)
                On Error GoTo 0
            End If
        Next
    End Sub

End Class
