<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmVfcMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents lstVBProjects As System.Windows.Forms.ListBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents chkIncludeCode As System.Windows.Forms.CheckBox
	Public WithEvents chkShowUnknown As System.Windows.Forms.CheckBox
	Public WithEvents txtGitRepoPath As System.Windows.Forms.TextBox
	Public WithEvents cmdBrowseRepo As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cbDeleteIgnoreFiles As System.Windows.Forms.CheckBox
	Public WithEvents fraOptions As System.Windows.Forms.GroupBox
	Public WithEvents cmdConvert As System.Windows.Forms.Button
	Public WithEvents lblProject As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmVfcMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.lstVBProjects = New System.Windows.Forms.ListBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.fraOptions = New System.Windows.Forms.GroupBox
		Me.chkIncludeCode = New System.Windows.Forms.CheckBox
		Me.chkShowUnknown = New System.Windows.Forms.CheckBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtGitRepoPath = New System.Windows.Forms.TextBox
		Me.cmdBrowseRepo = New System.Windows.Forms.Button
		Me.cbDeleteIgnoreFiles = New System.Windows.Forms.CheckBox
		Me.cmdConvert = New System.Windows.Forms.Button
		Me.lblProject = New System.Windows.Forms.Label
		Me.fraOptions.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(352, 316)
		Me.Location = New System.Drawing.Point(3, 18)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Tag = "VFC"
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmVfcMain"
		Me.lstVBProjects.Size = New System.Drawing.Size(336, 91)
		Me.lstVBProjects.IntegralHeight = False
		Me.lstVBProjects.Location = New System.Drawing.Point(8, 24)
		Me.lstVBProjects.TabIndex = 0
		Me.lstVBProjects.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstVBProjects.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstVBProjects.BackColor = System.Drawing.SystemColors.Window
		Me.lstVBProjects.CausesValidation = True
		Me.lstVBProjects.Enabled = True
		Me.lstVBProjects.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstVBProjects.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstVBProjects.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstVBProjects.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstVBProjects.Sorted = False
		Me.lstVBProjects.TabStop = True
		Me.lstVBProjects.Visible = True
		Me.lstVBProjects.MultiColumn = True
		Me.lstVBProjects.ColumnWidth = 168
		Me.lstVBProjects.Name = "lstVBProjects"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Exit"
		Me.cmdCancel.Size = New System.Drawing.Size(64, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(232, 272)
		Me.cmdCancel.TabIndex = 2
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.fraOptions.Text = "Convert Forms"
		Me.fraOptions.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraOptions.Size = New System.Drawing.Size(336, 144)
		Me.fraOptions.Location = New System.Drawing.Point(8, 120)
		Me.fraOptions.TabIndex = 3
		Me.fraOptions.BackColor = System.Drawing.SystemColors.Control
		Me.fraOptions.Enabled = True
		Me.fraOptions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraOptions.Visible = True
		Me.fraOptions.Padding = New System.Windows.Forms.Padding(0)
		Me.fraOptions.Name = "fraOptions"
		Me.chkIncludeCode.Text = "Include Code in UserForm Export"
		Me.chkIncludeCode.Size = New System.Drawing.Size(314, 20)
		Me.chkIncludeCode.Location = New System.Drawing.Point(10, 16)
		Me.chkIncludeCode.TabIndex = 5
		Me.chkIncludeCode.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkIncludeCode.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIncludeCode.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIncludeCode.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIncludeCode.BackColor = System.Drawing.SystemColors.Control
		Me.chkIncludeCode.CausesValidation = True
		Me.chkIncludeCode.Enabled = True
		Me.chkIncludeCode.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkIncludeCode.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIncludeCode.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIncludeCode.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIncludeCode.TabStop = True
		Me.chkIncludeCode.Visible = True
		Me.chkIncludeCode.Name = "chkIncludeCode"
		Me.chkShowUnknown.Text = "Show Placeholders for Unknown Controls"
		Me.chkShowUnknown.Size = New System.Drawing.Size(322, 20)
		Me.chkShowUnknown.Location = New System.Drawing.Point(10, 40)
		Me.chkShowUnknown.TabIndex = 6
		Me.chkShowUnknown.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkShowUnknown.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShowUnknown.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShowUnknown.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShowUnknown.BackColor = System.Drawing.SystemColors.Control
		Me.chkShowUnknown.CausesValidation = True
		Me.chkShowUnknown.Enabled = True
		Me.chkShowUnknown.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkShowUnknown.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShowUnknown.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShowUnknown.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShowUnknown.TabStop = True
		Me.chkShowUnknown.Visible = True
		Me.chkShowUnknown.Name = "chkShowUnknown"
		Me.Frame1.Text = "Git Repository Root"
		Me.Frame1.Size = New System.Drawing.Size(320, 48)
		Me.Frame1.Location = New System.Drawing.Point(10, 88)
		Me.Frame1.TabIndex = 7
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.txtGitRepoPath.AutoSize = False
		Me.txtGitRepoPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtGitRepoPath.Size = New System.Drawing.Size(240, 24)
		Me.txtGitRepoPath.Location = New System.Drawing.Point(10, 16)
		Me.txtGitRepoPath.ReadOnly = True
		Me.txtGitRepoPath.TabIndex = 8
		Me.txtGitRepoPath.TabStop = False
		Me.txtGitRepoPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtGitRepoPath.AcceptsReturn = True
		Me.txtGitRepoPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtGitRepoPath.CausesValidation = True
		Me.txtGitRepoPath.Enabled = True
		Me.txtGitRepoPath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtGitRepoPath.HideSelection = True
		Me.txtGitRepoPath.Maxlength = 0
		Me.txtGitRepoPath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtGitRepoPath.MultiLine = False
		Me.txtGitRepoPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtGitRepoPath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtGitRepoPath.Visible = True
		Me.txtGitRepoPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtGitRepoPath.Name = "txtGitRepoPath"
		Me.cmdBrowseRepo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBrowseRepo.Text = "Browse"
		Me.cmdBrowseRepo.Size = New System.Drawing.Size(64, 24)
		Me.cmdBrowseRepo.Location = New System.Drawing.Point(250, 16)
		Me.cmdBrowseRepo.TabIndex = 9
		Me.cmdBrowseRepo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseRepo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseRepo.CausesValidation = True
		Me.cmdBrowseRepo.Enabled = True
		Me.cmdBrowseRepo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseRepo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseRepo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseRepo.TabStop = True
		Me.cmdBrowseRepo.Name = "cmdBrowseRepo"
		Me.cbDeleteIgnoreFiles.Text = "Delete files in /ignore/ after processing"
		Me.cbDeleteIgnoreFiles.Size = New System.Drawing.Size(322, 20)
		Me.cbDeleteIgnoreFiles.Location = New System.Drawing.Point(10, 64)
		Me.cbDeleteIgnoreFiles.TabIndex = 10
		Me.cbDeleteIgnoreFiles.CheckState = System.Windows.Forms.CheckState.Checked
		Me.cbDeleteIgnoreFiles.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbDeleteIgnoreFiles.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.cbDeleteIgnoreFiles.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.cbDeleteIgnoreFiles.BackColor = System.Drawing.SystemColors.Control
		Me.cbDeleteIgnoreFiles.CausesValidation = True
		Me.cbDeleteIgnoreFiles.Enabled = True
		Me.cbDeleteIgnoreFiles.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cbDeleteIgnoreFiles.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbDeleteIgnoreFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbDeleteIgnoreFiles.Appearance = System.Windows.Forms.Appearance.Normal
		Me.cbDeleteIgnoreFiles.TabStop = True
		Me.cbDeleteIgnoreFiles.Visible = True
		Me.cbDeleteIgnoreFiles.Name = "cbDeleteIgnoreFiles"
		Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdConvert.Text = "Export Project for Source Control"
		Me.cmdConvert.Enabled = False
		Me.cmdConvert.Size = New System.Drawing.Size(184, 25)
		Me.cmdConvert.Location = New System.Drawing.Point(40, 272)
		Me.cmdConvert.TabIndex = 4
		Me.cmdConvert.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdConvert.BackColor = System.Drawing.SystemColors.Control
		Me.cmdConvert.CausesValidation = True
		Me.cmdConvert.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdConvert.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdConvert.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdConvert.TabStop = True
		Me.cmdConvert.Name = "cmdConvert"
		Me.lblProject.Text = "Select Project:"
		Me.lblProject.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProject.Size = New System.Drawing.Size(192, 16)
		Me.lblProject.Location = New System.Drawing.Point(16, 8)
		Me.lblProject.TabIndex = 1
		Me.lblProject.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblProject.BackColor = System.Drawing.SystemColors.Control
		Me.lblProject.Enabled = True
		Me.lblProject.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProject.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProject.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProject.UseMnemonic = True
		Me.lblProject.Visible = True
		Me.lblProject.AutoSize = False
		Me.lblProject.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProject.Name = "lblProject"
		Me.Controls.Add(lstVBProjects)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(fraOptions)
		Me.Controls.Add(cmdConvert)
		Me.Controls.Add(lblProject)
		Me.fraOptions.Controls.Add(chkIncludeCode)
		Me.fraOptions.Controls.Add(chkShowUnknown)
		Me.fraOptions.Controls.Add(Frame1)
		Me.fraOptions.Controls.Add(cbDeleteIgnoreFiles)
		Me.Frame1.Controls.Add(txtGitRepoPath)
		Me.Frame1.Controls.Add(cmdBrowseRepo)
		Me.fraOptions.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class