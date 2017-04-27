using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace VBA2VBProject
{
	internal partial class frmVfcMain : System.Windows.Forms.Form
	{


		public int SelectedProjectIndex {
			get {
				int functionReturnValue = 0;
				if (lstVBProjects.SelectedIndex > -1) {
					functionReturnValue = Convert.ToInt32(lstVBProjects.Text);
				}
				return functionReturnValue;
			}
		}

		public string SelectedProjectFilename {
			get {
				string functionReturnValue = null;
				if (lstVBProjects.SelectedIndex > -1) {
					functionReturnValue = lstVBProjects.List(lstVBProjects.SelectedIndex, 2);
				}
				return functionReturnValue;
			}
		}

		public string SelectedProjectGitDirectory {
			get {
				string functionReturnValue = null;
				if (lstVBProjects.SelectedIndex > -1) {
					functionReturnValue = txtGitRepoPath.Text + modFileFunctions.StripExtensionFromFileName(modFileFunctions.GetFileNameFromPath(lstVBProjects.SelectedItem(lstVBProjects.SelectedIndex, 2)));
					//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					if (string.IsNullOrEmpty(FileSystem.Dir(SelectedProjectGitDirectory, FileAttribute.Directory)))
						FileSystem.MkDir(SelectedProjectGitDirectory);
				}
				return functionReturnValue;
			}
		}

		public string SelectedProjectGitIgnoreDirectory {
			get {
				string functionReturnValue = null;
				if (lstVBProjects.SelectedIndex > -1) {
					functionReturnValue = SelectedProjectGitDirectory + "\\ignore\\";
					//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					if (string.IsNullOrEmpty(FileSystem.Dir(SelectedProjectGitIgnoreDirectory, FileAttribute.Directory)))
						FileSystem.MkDir(SelectedProjectGitIgnoreDirectory);
				}
				return functionReturnValue;
			}
		}


		public bool IncludeCode {
			get { return (chkIncludeCode.CheckState == true); }
		}

		public bool ShowUnknown {
			get { return (chkShowUnknown.CheckState == true); }
		}

		private void cmdBrowseRepo_Click(System.Object eventSender, System.EventArgs eventArgs)
		{

			txtGitRepoPath.Text = modFileFunctions.BrowseFolder(ref "Select GIT Repo Base Folder", ref Interaction.Environ("USERPROFILE") + "\\Source\\Repos\\Autocad Automation\\");
			//MsgBox Environ("USERPROFILE")
			//"C:\Users\acunningham\Source\Repos"
		}

		private void cmdCancel_Click(System.Object eventSender, System.EventArgs eventArgs)
		{
			this.Close();
		}

		private void cmdConvert_Click(System.Object eventSender, System.EventArgs eventArgs)
		{
			cmdCancel.Text = "Exit";
			modVfcMain.ProcessProject();
			modDevTools.ExportSourceFiles(ref SelectedProjectGitDirectory + "\\");
		}



//UPGRADE_WARNING: Event lstVBProjects.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
		private void lstVBProjects_SelectedIndexChanged(System.Object eventSender, System.EventArgs eventArgs)
		{
			Microsoft.Vbe.Interop.VBProject objProj = default(Microsoft.Vbe.Interop.VBProject);
			short intIndex = 0;

			if (lstVBProjects.SelectedIndex > -1) {
				//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				objProj = AutoCADAcadApplication_definst.Application.VBE.VBProjects(this.SelectedProjectIndex);
				cmdConvert.Enabled = true;

			}

			//UPGRADE_NOTE: Object objProj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objProj = null;
		}



		private void UserForm_Activate()
		{
			Microsoft.Vbe.Interop.VBProjects objProjs = default(Microsoft.Vbe.Interop.VBProjects);
			Microsoft.Vbe.Interop.VBProject objProj = default(Microsoft.Vbe.Interop.VBProject);
			short intIndex = 0;
			string strTemp = null;

			 // ERROR: Not supported in C#: OnErrorStatement


			//Load up list box with VBProjects
			//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objProjs = AutoCADAcadApplication_definst.Application.VBE.VBProjects;
			cmdConvert.Enabled = false;
			//  lstMSForms.Clear
			lstVBProjects.Items.Clear();
			//UPGRADE_WARNING: Couldn't resolve default property of object lstVBProjects.ColumnCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lstVBProjects.ColumnCount = 3;
			//UPGRADE_WARNING: Couldn't resolve default property of object lstVBProjects.ColumnWidths. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lstVBProjects.ColumnWidths = "20 pt;90 pt;400 pt";
			short listIndex = 0;
			listIndex = 1;
			for (intIndex = 1; intIndex <= objProjs.Count; intIndex++) {
				objProj = objProjs.Item(intIndex);
				//do not list the exporter macro
				if ((objProj.fileName != objProjs.VBE.ActiveVBProject.fileName)) {
					lstVBProjects.Items.Add(Convert.ToString(intIndex));
					lstVBProjects.List(listIndex - 1, 1) = objProj.Name;
					lstVBProjects.List(listIndex - 1, 2) = objProj.fileName;
					listIndex = listIndex + 1;
				}
			}
			txtGitRepoPath.Text = Interaction.Environ("USERPROFILE") + "\\Source\\Repos\\Autocad Automation\\";
			SubExit:


			//UPGRADE_NOTE: Object objProj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objProj = null;
			//UPGRADE_NOTE: Object objProjs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objProjs = null;
			return;
			ErrorHandler:
			switch (Err().Number) {
				case 76:
					//Path not found
					strTemp = "VFC can only work with SAVED projects," + Constants.vbCrLf + "please save all newly created projects" + Constants.vbCrLf + "and start again.";
					Interaction.MsgBox(strTemp, MsgBoxStyle.Critical, modVfcMain.ccAPPNAME);
					break;
				default:
					//Huh??
					strTemp = modVfcMain.ccAPPNAME + " ERROR" + Constants.vbCrLf + "VFC_001: Error in UserForm_Activate" + Constants.vbCrLf + "Description: " + Err().Description + Constants.vbCrLf + "Source: " + Err().Source + Constants.vbCrLf + "Number: " + Convert.ToString(Err().Number) + Constants.vbCrLf + "VFC Ver: " + modVfcMain.ccAPPVER;
					Interaction.MsgBox(strTemp, MsgBoxStyle.Critical, modVfcMain.ccAPPNAME);
					break;
			}
			this.Close();
			 // ERROR: Not supported in C#: ResumeStatement

		}

		private void UserForm_Terminate()
		{
			Debug.Print("UserForm_Terminate frmVfcMain");
		}
	}
}
