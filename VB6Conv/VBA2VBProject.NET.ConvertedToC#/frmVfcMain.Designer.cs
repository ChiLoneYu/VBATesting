using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
namespace VBA2VBProject
{
	[Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
	partial class frmVfcMain
	{
		#region "Windows Form Designer generated code "
		[System.Diagnostics.DebuggerNonUserCode()]
		public frmVfcMain() : base()
		{
			//This call is required by the Windows Form Designer.
			InitializeComponent();
		}
//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]
		protected override void Dispose(bool Disposing)
		{
			if (Disposing) {
				if ((components != null)) {
					components.Dispose();
				}
			}
			base.Dispose(Disposing);
		}
//Required by the Windows Form Designer
		private System.ComponentModel.IContainer components;
		public System.Windows.Forms.ToolTip ToolTip1;
		private System.Windows.Forms.ListBox withEventsField_lstVBProjects;
		public System.Windows.Forms.ListBox lstVBProjects {
			get { return withEventsField_lstVBProjects; }
			set {
				if (withEventsField_lstVBProjects != null) {
					withEventsField_lstVBProjects.SelectedIndexChanged -= lstVBProjects_SelectedIndexChanged;
				}
				withEventsField_lstVBProjects = value;
				if (withEventsField_lstVBProjects != null) {
					withEventsField_lstVBProjects.SelectedIndexChanged += lstVBProjects_SelectedIndexChanged;
				}
			}
		}
		private System.Windows.Forms.Button withEventsField_cmdCancel;
		public System.Windows.Forms.Button cmdCancel {
			get { return withEventsField_cmdCancel; }
			set {
				if (withEventsField_cmdCancel != null) {
					withEventsField_cmdCancel.Click -= cmdCancel_Click;
				}
				withEventsField_cmdCancel = value;
				if (withEventsField_cmdCancel != null) {
					withEventsField_cmdCancel.Click += cmdCancel_Click;
				}
			}
		}
		public System.Windows.Forms.CheckBox chkIncludeCode;
		public System.Windows.Forms.CheckBox chkShowUnknown;
		public System.Windows.Forms.TextBox txtGitRepoPath;
		private System.Windows.Forms.Button withEventsField_cmdBrowseRepo;
		public System.Windows.Forms.Button cmdBrowseRepo {
			get { return withEventsField_cmdBrowseRepo; }
			set {
				if (withEventsField_cmdBrowseRepo != null) {
					withEventsField_cmdBrowseRepo.Click -= cmdBrowseRepo_Click;
				}
				withEventsField_cmdBrowseRepo = value;
				if (withEventsField_cmdBrowseRepo != null) {
					withEventsField_cmdBrowseRepo.Click += cmdBrowseRepo_Click;
				}
			}
		}
		public System.Windows.Forms.GroupBox Frame1;
		public System.Windows.Forms.CheckBox cbDeleteIgnoreFiles;
		public System.Windows.Forms.GroupBox fraOptions;
		private System.Windows.Forms.Button withEventsField_cmdConvert;
		public System.Windows.Forms.Button cmdConvert {
			get { return withEventsField_cmdConvert; }
			set {
				if (withEventsField_cmdConvert != null) {
					withEventsField_cmdConvert.Click -= cmdConvert_Click;
				}
				withEventsField_cmdConvert = value;
				if (withEventsField_cmdConvert != null) {
					withEventsField_cmdConvert.Click += cmdConvert_Click;
				}
			}
		}
		public System.Windows.Forms.Label lblProject;
//NOTE: The following procedure is required by the Windows Form Designer
//It can be modified using the Windows Form Designer.
//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmVfcMain));
			this.components = new System.ComponentModel.Container();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(components);
			this.lstVBProjects = new System.Windows.Forms.ListBox();
			this.cmdCancel = new System.Windows.Forms.Button();
			this.fraOptions = new System.Windows.Forms.GroupBox();
			this.chkIncludeCode = new System.Windows.Forms.CheckBox();
			this.chkShowUnknown = new System.Windows.Forms.CheckBox();
			this.Frame1 = new System.Windows.Forms.GroupBox();
			this.txtGitRepoPath = new System.Windows.Forms.TextBox();
			this.cmdBrowseRepo = new System.Windows.Forms.Button();
			this.cbDeleteIgnoreFiles = new System.Windows.Forms.CheckBox();
			this.cmdConvert = new System.Windows.Forms.Button();
			this.lblProject = new System.Windows.Forms.Label();
			this.fraOptions.SuspendLayout();
			this.Frame1.SuspendLayout();
			this.SuspendLayout();
			this.ToolTip1.Active = true;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.ClientSize = new System.Drawing.Size(352, 316);
			this.Location = new System.Drawing.Point(3, 18);
			this.Font = new System.Drawing.Font("Tahoma", 8.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Tag = "VFC";
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ControlBox = true;
			this.Enabled = true;
			this.KeyPreview = false;
			this.Cursor = System.Windows.Forms.Cursors.Default;
			this.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ShowInTaskbar = true;
			this.HelpButton = false;
			this.WindowState = System.Windows.Forms.FormWindowState.Normal;
			this.Name = "frmVfcMain";
			this.lstVBProjects.Size = new System.Drawing.Size(336, 91);
			this.lstVBProjects.IntegralHeight = false;
			this.lstVBProjects.Location = new System.Drawing.Point(8, 24);
			this.lstVBProjects.TabIndex = 0;
			this.lstVBProjects.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.lstVBProjects.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lstVBProjects.BackColor = System.Drawing.SystemColors.Window;
			this.lstVBProjects.CausesValidation = true;
			this.lstVBProjects.Enabled = true;
			this.lstVBProjects.ForeColor = System.Drawing.SystemColors.WindowText;
			this.lstVBProjects.Cursor = System.Windows.Forms.Cursors.Default;
			this.lstVBProjects.SelectionMode = System.Windows.Forms.SelectionMode.One;
			this.lstVBProjects.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lstVBProjects.Sorted = false;
			this.lstVBProjects.TabStop = true;
			this.lstVBProjects.Visible = true;
			this.lstVBProjects.MultiColumn = true;
			this.lstVBProjects.ColumnWidth = 168;
			this.lstVBProjects.Name = "lstVBProjects";
			this.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.CancelButton = this.cmdCancel;
			this.cmdCancel.Text = "Exit";
			this.cmdCancel.Size = new System.Drawing.Size(64, 25);
			this.cmdCancel.Location = new System.Drawing.Point(232, 272);
			this.cmdCancel.TabIndex = 2;
			this.cmdCancel.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.cmdCancel.BackColor = System.Drawing.SystemColors.Control;
			this.cmdCancel.CausesValidation = true;
			this.cmdCancel.Enabled = true;
			this.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText;
			this.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cmdCancel.TabStop = true;
			this.cmdCancel.Name = "cmdCancel";
			this.fraOptions.Text = "Convert Forms";
			this.fraOptions.Font = new System.Drawing.Font("Tahoma", 8.25f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.fraOptions.Size = new System.Drawing.Size(336, 144);
			this.fraOptions.Location = new System.Drawing.Point(8, 120);
			this.fraOptions.TabIndex = 3;
			this.fraOptions.BackColor = System.Drawing.SystemColors.Control;
			this.fraOptions.Enabled = true;
			this.fraOptions.ForeColor = System.Drawing.SystemColors.ControlText;
			this.fraOptions.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.fraOptions.Visible = true;
			this.fraOptions.Padding = new System.Windows.Forms.Padding(0);
			this.fraOptions.Name = "fraOptions";
			this.chkIncludeCode.Text = "Include Code in UserForm Export";
			this.chkIncludeCode.Size = new System.Drawing.Size(314, 20);
			this.chkIncludeCode.Location = new System.Drawing.Point(10, 16);
			this.chkIncludeCode.TabIndex = 5;
			this.chkIncludeCode.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkIncludeCode.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.chkIncludeCode.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.chkIncludeCode.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
			this.chkIncludeCode.BackColor = System.Drawing.SystemColors.Control;
			this.chkIncludeCode.CausesValidation = true;
			this.chkIncludeCode.Enabled = true;
			this.chkIncludeCode.ForeColor = System.Drawing.SystemColors.ControlText;
			this.chkIncludeCode.Cursor = System.Windows.Forms.Cursors.Default;
			this.chkIncludeCode.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.chkIncludeCode.Appearance = System.Windows.Forms.Appearance.Normal;
			this.chkIncludeCode.TabStop = true;
			this.chkIncludeCode.Visible = true;
			this.chkIncludeCode.Name = "chkIncludeCode";
			this.chkShowUnknown.Text = "Show Placeholders for Unknown Controls";
			this.chkShowUnknown.Size = new System.Drawing.Size(322, 20);
			this.chkShowUnknown.Location = new System.Drawing.Point(10, 40);
			this.chkShowUnknown.TabIndex = 6;
			this.chkShowUnknown.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkShowUnknown.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.chkShowUnknown.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.chkShowUnknown.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
			this.chkShowUnknown.BackColor = System.Drawing.SystemColors.Control;
			this.chkShowUnknown.CausesValidation = true;
			this.chkShowUnknown.Enabled = true;
			this.chkShowUnknown.ForeColor = System.Drawing.SystemColors.ControlText;
			this.chkShowUnknown.Cursor = System.Windows.Forms.Cursors.Default;
			this.chkShowUnknown.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.chkShowUnknown.Appearance = System.Windows.Forms.Appearance.Normal;
			this.chkShowUnknown.TabStop = true;
			this.chkShowUnknown.Visible = true;
			this.chkShowUnknown.Name = "chkShowUnknown";
			this.Frame1.Text = "Git Repository Root";
			this.Frame1.Size = new System.Drawing.Size(320, 48);
			this.Frame1.Location = new System.Drawing.Point(10, 88);
			this.Frame1.TabIndex = 7;
			this.Frame1.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.Frame1.BackColor = System.Drawing.SystemColors.Control;
			this.Frame1.Enabled = true;
			this.Frame1.ForeColor = System.Drawing.SystemColors.ControlText;
			this.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.Frame1.Visible = true;
			this.Frame1.Padding = new System.Windows.Forms.Padding(0);
			this.Frame1.Name = "Frame1";
			this.txtGitRepoPath.AutoSize = false;
			this.txtGitRepoPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtGitRepoPath.Size = new System.Drawing.Size(240, 24);
			this.txtGitRepoPath.Location = new System.Drawing.Point(10, 16);
			this.txtGitRepoPath.ReadOnly = true;
			this.txtGitRepoPath.TabIndex = 8;
			this.txtGitRepoPath.TabStop = false;
			this.txtGitRepoPath.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.txtGitRepoPath.AcceptsReturn = true;
			this.txtGitRepoPath.BackColor = System.Drawing.SystemColors.Window;
			this.txtGitRepoPath.CausesValidation = true;
			this.txtGitRepoPath.Enabled = true;
			this.txtGitRepoPath.ForeColor = System.Drawing.SystemColors.WindowText;
			this.txtGitRepoPath.HideSelection = true;
			this.txtGitRepoPath.MaxLength = 0;
			this.txtGitRepoPath.Cursor = System.Windows.Forms.Cursors.IBeam;
			this.txtGitRepoPath.Multiline = false;
			this.txtGitRepoPath.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtGitRepoPath.ScrollBars = System.Windows.Forms.ScrollBars.None;
			this.txtGitRepoPath.Visible = true;
			this.txtGitRepoPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.txtGitRepoPath.Name = "txtGitRepoPath";
			this.cmdBrowseRepo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.cmdBrowseRepo.Text = "Browse";
			this.cmdBrowseRepo.Size = new System.Drawing.Size(64, 24);
			this.cmdBrowseRepo.Location = new System.Drawing.Point(250, 16);
			this.cmdBrowseRepo.TabIndex = 9;
			this.cmdBrowseRepo.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.cmdBrowseRepo.BackColor = System.Drawing.SystemColors.Control;
			this.cmdBrowseRepo.CausesValidation = true;
			this.cmdBrowseRepo.Enabled = true;
			this.cmdBrowseRepo.ForeColor = System.Drawing.SystemColors.ControlText;
			this.cmdBrowseRepo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdBrowseRepo.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cmdBrowseRepo.TabStop = true;
			this.cmdBrowseRepo.Name = "cmdBrowseRepo";
			this.cbDeleteIgnoreFiles.Text = "Delete files in /ignore/ after processing";
			this.cbDeleteIgnoreFiles.Size = new System.Drawing.Size(322, 20);
			this.cbDeleteIgnoreFiles.Location = new System.Drawing.Point(10, 64);
			this.cbDeleteIgnoreFiles.TabIndex = 10;
			this.cbDeleteIgnoreFiles.CheckState = System.Windows.Forms.CheckState.Checked;
			this.cbDeleteIgnoreFiles.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.cbDeleteIgnoreFiles.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.cbDeleteIgnoreFiles.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
			this.cbDeleteIgnoreFiles.BackColor = System.Drawing.SystemColors.Control;
			this.cbDeleteIgnoreFiles.CausesValidation = true;
			this.cbDeleteIgnoreFiles.Enabled = true;
			this.cbDeleteIgnoreFiles.ForeColor = System.Drawing.SystemColors.ControlText;
			this.cbDeleteIgnoreFiles.Cursor = System.Windows.Forms.Cursors.Default;
			this.cbDeleteIgnoreFiles.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cbDeleteIgnoreFiles.Appearance = System.Windows.Forms.Appearance.Normal;
			this.cbDeleteIgnoreFiles.TabStop = true;
			this.cbDeleteIgnoreFiles.Visible = true;
			this.cbDeleteIgnoreFiles.Name = "cbDeleteIgnoreFiles";
			this.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.cmdConvert.Text = "Export Project for Source Control";
			this.cmdConvert.Enabled = false;
			this.cmdConvert.Size = new System.Drawing.Size(184, 25);
			this.cmdConvert.Location = new System.Drawing.Point(40, 272);
			this.cmdConvert.TabIndex = 4;
			this.cmdConvert.Font = new System.Drawing.Font("Arial", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.cmdConvert.BackColor = System.Drawing.SystemColors.Control;
			this.cmdConvert.CausesValidation = true;
			this.cmdConvert.ForeColor = System.Drawing.SystemColors.ControlText;
			this.cmdConvert.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdConvert.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cmdConvert.TabStop = true;
			this.cmdConvert.Name = "cmdConvert";
			this.lblProject.Text = "Select Project:";
			this.lblProject.Font = new System.Drawing.Font("Tahoma", 8.25f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, Convert.ToByte(0));
			this.lblProject.Size = new System.Drawing.Size(192, 16);
			this.lblProject.Location = new System.Drawing.Point(16, 8);
			this.lblProject.TabIndex = 1;
			this.lblProject.TextAlign = System.Drawing.ContentAlignment.TopLeft;
			this.lblProject.BackColor = System.Drawing.SystemColors.Control;
			this.lblProject.Enabled = true;
			this.lblProject.ForeColor = System.Drawing.SystemColors.ControlText;
			this.lblProject.Cursor = System.Windows.Forms.Cursors.Default;
			this.lblProject.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lblProject.UseMnemonic = true;
			this.lblProject.Visible = true;
			this.lblProject.AutoSize = false;
			this.lblProject.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.lblProject.Name = "lblProject";
			this.Controls.Add(lstVBProjects);
			this.Controls.Add(cmdCancel);
			this.Controls.Add(fraOptions);
			this.Controls.Add(cmdConvert);
			this.Controls.Add(lblProject);
			this.fraOptions.Controls.Add(chkIncludeCode);
			this.fraOptions.Controls.Add(chkShowUnknown);
			this.fraOptions.Controls.Add(Frame1);
			this.fraOptions.Controls.Add(cbDeleteIgnoreFiles);
			this.Frame1.Controls.Add(txtGitRepoPath);
			this.Frame1.Controls.Add(cmdBrowseRepo);
			this.fraOptions.ResumeLayout(false);
			this.Frame1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();
		}
		#endregion
	}
}
