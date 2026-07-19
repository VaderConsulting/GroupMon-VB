<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatus
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStatus))
        Me.ButtonQuery = New System.Windows.Forms.Button
        Me.lstResults = New System.Windows.Forms.ListBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'ButtonQuery
        '
        Me.ButtonQuery.Location = New System.Drawing.Point(12, 12)
        Me.ButtonQuery.Name = "ButtonQuery"
        Me.ButtonQuery.Size = New System.Drawing.Size(75, 23)
        Me.ButtonQuery.TabIndex = 0
        Me.ButtonQuery.Text = "Query"
        Me.ButtonQuery.UseVisualStyleBackColor = True
        '
        'lstResults
        '
        Me.lstResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Location = New System.Drawing.Point(12, 41)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(260, 212)
        Me.lstResults.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(96, 12)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 13)
        Me.lblStatus.TabIndex = 2
        '
        'frmStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 264)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.ButtonQuery)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(300, 300)
        Me.Name = "frmStatus"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Group Status"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonQuery As System.Windows.Forms.Button
    Friend WithEvents lstResults As System.Windows.Forms.ListBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label

End Class
