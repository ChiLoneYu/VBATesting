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
	static class modVfcMain
	{

		public const string ccAPPNAME = "VBA²VB Form Converter";
		public const string ccAPPVER = "0.12";

		private static Microsoft.Vbe.Interop.VBComponent mobjMSForm;
		private static Microsoft.Vbe.Interop.Forms.Control mobjControl;

		private static string mstrSaveAs;
		private static bool mblnIncludeCode;
		private static bool mblnShowUnknown;
		private static short mintIndent;
		private static short mintContainer;
		private static short mintFile;
		private static short mintUnknownControls;

		public static short mstrAcadVersion;
		public static short mstrAcadPlatform;
		public static string mstrAcadVertical;

		//As Object
		public static void SelectFormToConvert()
		{
			//START HERE
			My.MyProject.Forms.frmVfcMain.ShowDialog();
		}

		public static void ProcessForm(ref string strFormName, ref bool blnProceed)
		{
			Microsoft.Vbe.Interop.Forms.Control objControl = default(Microsoft.Vbe.Interop.Forms.Control);
			string strMessage = null;

			//Setup variables
			//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjMSForm = AutoCADAcadApplication_definst.Application.VBE.VBProjects(My.MyProject.Forms.frmVfcMain.SelectedProjectIndex).VBComponents(strFormName);
			mstrSaveAs = My.MyProject.Forms.frmVfcMain.SelectedProjectGitIgnoreDirectory + mobjMSForm.Name + ".frmcmp";
			mblnIncludeCode = My.MyProject.Forms.frmVfcMain.IncludeCode;
			mblnShowUnknown = My.MyProject.Forms.frmVfcMain.ShowUnknown;
			mintIndent = 0;
			mintContainer = 0;
			mintFile = FreeFile();
			mintUnknownControls = 0;

			//Debug report
			Debug.Print("@--- VFC Debug Report ---@");
			Debug.Print("  Form Being Exported: " + mobjMSForm.Name);
			Debug.Print("  Filename of New Form: " + mstrSaveAs);
			Debug.Print("  Including Code?: " + mblnIncludeCode);
			Debug.Print("  Showing Unknown Controls?: " + mblnShowUnknown);

			//Remove the following IF block if you
			//dont want it to check for overwriting
			//  If Dir(mstrSaveAs) <> "" Then
			//    If MsgBox(mstrSaveAs & " already exists." & vbCrLf & "Do you want to replace it?", _
			//'              vbYesNo Or vbExclamation, ccAPPNAME) = vbNo Then
			//      blnProceed = False
			//    End If
			//  End If

			if (blnProceed) {
				//Convert form
				FileSystem.FileOpen(mintFile, mstrSaveAs, OpenMode.Output);
				WriteFormHeader();
				WriteFormProperties();
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjMSForm.Designer.Controls. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				foreach ( objControl in mobjMSForm.Designer.Controls) {
					//UPGRADE_WARNING: Couldn't resolve default property of object objControl.Parent.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					if (objControl.Parent.Name == mobjMSForm.Name) {
						mobjControl = objControl;
						ProcessControl(ref 0);
					}
				}
				WriteFormFooter();
				WriteFormCode();
				FileSystem.FileClose(mintFile);

				//Completion message
				strMessage = "Form Name = " + strFormName + Constants.vbCrLf + "Form Conversion completed ";
				if (mintUnknownControls > 0) {
					strMessage = strMessage + "with errors." + Constants.vbCrLf;
					strMessage = strMessage + Convert.ToString(mintUnknownControls) + " unknown controls were found." + Constants.vbCrLf;
					if (mblnShowUnknown) {
						strMessage = strMessage + "Please examine the form before use." + Constants.vbCrLf;
					} else {
						strMessage = strMessage + "They have been omitted from the form." + Constants.vbCrLf;
					}
				} else {
					strMessage = strMessage + "successfully." + Constants.vbCrLf;
				}
				strMessage = strMessage + "The new form file was saved to:" + Constants.vbCrLf + Constants.vbCrLf + mstrSaveAs;
				Debug.Print(strMessage);
				//, vbInformation, ccAPPNAME

			}

			//UPGRADE_NOTE: Object mobjControl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjControl = null;
			//UPGRADE_NOTE: Object objControl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objControl = null;
			//UPGRADE_NOTE: Object mobjMSForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjMSForm = null;
		}

		private static void ProcessControl(ref short intContainer)
		{
			//Debug.Print mobjControl.Name, mobjControl.Parent.Name
			mintContainer = intContainer;

			//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			if (mobjControl is Microsoft.Vbe.Interop.Forms.Label) {
				WriteLabelProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.TextBox) {
				WriteTextBoxProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.CheckBox) {
				WriteCheckBoxProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.OptionButton) {
				WriteOptionButtonProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.CommandButton) {
				WriteCommandButtonProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.ToggleButton) {
				WriteToggleButtonProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.Image) {
				WriteImageProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.ListBox) {
				WriteListBoxProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.ComboBox) {
				WriteComboBoxProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.ScrollBar) {
				WriteScrollBarProperties();
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			} else if (mobjControl is Microsoft.Vbe.Interop.Forms.Frame) {
				WriteFrameProperties();
			} else {
				WriteUnknownProperties();
			}

		}

		private static void WriteFrameProperties()
		{
			Microsoft.Vbe.Interop.Forms.Control objControl = default(Microsoft.Vbe.Interop.Forms.Control);
			string strCurrentParent = null;

			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.Frame " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			//Write control properties
			WriteBackColor(ref 0x8000000f);
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000012);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteTabIndex();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			strCurrentParent = mobjControl.Name;
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjMSForm.Designer.Controls. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			foreach ( objControl in mobjMSForm.Designer.Controls) {
				//UPGRADE_WARNING: Couldn't resolve default property of object objControl.Parent.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if (objControl.Parent.Name == strCurrentParent) {
					mobjControl = objControl;
					ProcessControl(ref 1);
				}
			}

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
			//UPGRADE_NOTE: Object objControl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objControl = null;
		}

		private static void WriteLabelProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.Label " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteAlignmentLabel();
			WriteAutoSize();
			WriteBackColor(ref 0x8000000f);
			WriteBackStyle();
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000012);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteTabIndex();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteTextBoxProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.TextBox " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteAlignmentLabel();
			WriteBackColor(ref 0x80000005);
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000008);
			WriteHeight();
			WriteHelpContextID();
			WriteHideSelection();
			WriteLeft();
			WriteLocked();
			WriteMaxLength();
			WriteMousePointer();
			WriteMultiLine();
			WritePasswordChar();
			WriteScrollBars();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteText();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteCheckBoxProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.CheckBox " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteAlignmentCheckBox();
			WriteBackColor(ref 0x8000000f);
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000012);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteValueCheckBox();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteOptionButtonProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.OptionButton " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteAlignmentCheckBox();
			WriteBackColor(ref 0x8000000f);
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000012);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteValueOptionButton();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteCommandButtonProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.CommandButton " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteBackColor(ref 0x8000000f);
			WriteCancel();
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteStyleCommandButton();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteToggleButtonProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.CheckBox " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteBackColor(ref 0x8000000f);
			WriteCaption();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000012);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteMousePointer();
			WriteStyleToggleButton();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteValueCheckBox();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteImageProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.Image " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteEnabled();
			WriteHeight();
			WriteLeft();
			WriteMousePointer();
			WriteStretch();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteListBoxProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.ListBox " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteBackColor(ref 0x80000005);
			WriteColumns();
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000008);
			WriteHeight();
			WriteHelpContextID();
			WriteIntegralHeight();
			WriteLeft();
			WriteMousePointer();
			WriteMultiSelect();
			WriteStyleListBox();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteComboBoxProperties()
		{
			FileSystem.PrintLine(mintFile, Indent() + "Begin VB.ComboBox " + mobjControl.Name);
			mintIndent = mintIndent + 1;

			WriteBackColor(ref 0x80000005);
			WriteEnabled();
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (ParentFontDifferent())
				WriteFont(ref mobjControl.Font);
			WriteForeColor(ref 0x80000008);
			WriteHeight();
			WriteHelpContextID();
			WriteLeft();
			WriteLocked();
			WriteMousePointer();
			WriteStyleComboBox();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteText();
			WriteToolTipText();
			WriteTop();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteScrollBarProperties()
		{
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Orientation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Orientation == Microsoft.Vbe.Interop.Forms.fmOrientation.fmOrientationAuto) {
				if (mobjControl.Width > mobjControl.Height) {
					FileSystem.PrintLine(mintFile, Indent() + "Begin VB.HScrollBar " + mobjControl.Name);
				} else {
					FileSystem.PrintLine(mintFile, Indent() + "Begin VB.VScrollBar " + mobjControl.Name);
				}
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Orientation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Orientation == Microsoft.Vbe.Interop.Forms.fmOrientation.fmOrientationVertical) {
				FileSystem.PrintLine(mintFile, Indent() + "Begin VB.VScrollBar " + mobjControl.Name);
			//If mobjControl.Orientation = fmOrientationHorizontal Then
			} else {
				FileSystem.PrintLine(mintFile, Indent() + "Begin VB.HScrollBar " + mobjControl.Name);
			}
			mintIndent = mintIndent + 1;

			WriteEnabled();
			WriteHeight();
			WriteHelpContextID();
			WriteLargeChange();
			WriteLeft();
			WriteMaxScrollBar();
			WriteMinScrollBar();
			WriteMousePointer();
			WriteSmallChange();
			WriteTabIndex();
			WriteTabStop();
			WriteTag();
			WriteTop();
			WriteValueScrollBar();
			WriteVisible();
			WriteWidth();

			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "End");
		}

		private static void WriteUnknownProperties()
		{

			mintUnknownControls = mintUnknownControls + 1;

			if (mblnShowUnknown) {
				//Show the unknown control on the converted form as a red label
				FileSystem.PrintLine(mintFile, Indent() + "Begin VB.Label " + mobjControl.Name);
				mintIndent = mintIndent + 1;

				FileSystem.PrintLine(mintFile, FormatProperty(ref "Alignment") + "2  'Center");
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Appearance") + "0  'Flat");
				FileSystem.PrintLine(mintFile, FormatProperty(ref "BackColor") + "&H000000FF&");
				FileSystem.PrintLine(mintFile, FormatProperty(ref "BorderStyle") + "1  'Fixed Single");
				//UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Caption") + FormatString(ref mobjControl.Name + " - " + Information.TypeName(mobjControl)));
				WriteHeight();
				WriteLeft();
				WriteTop();
				WriteWidth();

				mintIndent = mintIndent - 1;
				FileSystem.PrintLine(mintFile, Indent() + "End");
			}

			//Print the unknown control to the immediate window
			Debug.Print("@--- Unknown Control Found ---@");
			Debug.Print("  Control Name: " + mobjControl.Name);
			//UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Debug.Print("  Control Type: " + Information.TypeName(mobjControl));

		}

		private static void WriteAlignmentLabel()
		{
			//ALIGNMENT
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.TextAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.TextAlign != Microsoft.Vbe.Interop.Forms.fmTextAlign.fmTextAlignLeft) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.TextAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if (mobjControl.TextAlign == Microsoft.Vbe.Interop.Forms.fmTextAlign.fmTextAlignRight) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "Alignment") + "1  'Right Jusify");
					//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.TextAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				} else if (mobjControl.TextAlign == Microsoft.Vbe.Interop.Forms.fmTextAlign.fmTextAlignCenter) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "Alignment") + "2  'Center");
				}
			}
		}

		private static void WriteAlignmentCheckBox()
		{
			//ALIGNMENT
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Alignment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Alignment != Microsoft.Vbe.Interop.Forms.fmAlignment.fmAlignmentRight) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Alignment") + "1  'Right Jusify");
			}
		}

		private static void WriteAutoSize()
		{
			//AUTOSIZE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.AutoSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.AutoSize == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "AutoSize") + "-1  'True");
			}
		}

		private static void WriteBackColor(ref int lngDefault)
		{
			//BACKCOLOR
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.BackColor != lngDefault) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "BackColor") + FormatHex(ref mobjControl.BackColor));
			}
		}

		private static void WriteBackStyle()
		{
			//BACKSTYLE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.BackStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.BackStyle != Microsoft.Vbe.Interop.Forms.fmBackStyle.fmBackStyleOpaque) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "BackStyle") + "0  'Transparent");
			}
		}

		private static void WriteCancel()
		{
			//CANCEL
			if (mobjControl.Cancel == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Cancel") + "-1  'True");
			}
		}

		private static void WriteCaption()
		{
			string strCaption = null;
			string strChar = null;
			string strTemp = null;
			short intPos = 0;

			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strCaption = mobjControl.Caption;
			if (!string.IsNullOrEmpty(strCaption)) {
				//Find &'s and replace with &&'s
				for (intPos = 1; intPos <= Strings.Len(strCaption); intPos++) {
					strChar = Strings.Mid(strCaption, intPos, 1);
					if (strChar == "&") {
						strTemp = strTemp + "&";
					}
					strTemp = strTemp + strChar;
				}
				//Add mnemonic
				//UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				if (mobjControl is Microsoft.Vbe.Interop.Forms.Label | mobjControl is Microsoft.Vbe.Interop.Forms.CheckBox | mobjControl is Microsoft.Vbe.Interop.Forms.OptionButton | mobjControl is Microsoft.Vbe.Interop.Forms.CommandButton | mobjControl is Microsoft.Vbe.Interop.Forms.ToggleButton) {
					//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Accelerator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strChar = mobjControl.Accelerator;
					if (!string.IsNullOrEmpty(strChar)) {
						intPos = Strings.InStr(1, strTemp, strChar, CompareMethod.Binary);
						if (intPos > 0) {
							strTemp = Strings.Left(strTemp, intPos - 1) + "&" + Strings.Mid(strTemp, intPos);
						}
					}
				}
				//CAPTION
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Caption") + FormatString(ref strTemp));
			}
		}

		private static void WriteColumns()
		{
			//COLUMNS
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ColumnCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.ColumnCount > 1) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ColumnCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Columns") + Convert.ToString(mobjControl.ColumnCount - 1));
			}
		}

		private static void WriteDefault()
		{
			//DEFAULT
			if (mobjControl.Default == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Default") + "-1  'True");
			}
		}

		private static void WriteEnabled()
		{
			//ENABLED
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Enabled == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Enabled") + "0  'False");
			}
		}

		private static void WriteForeColor(ref int lngDefault)
		{
			//FORECOLOR
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.ForeColor != lngDefault) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "ForeColor") + FormatHex(ref mobjControl.ForeColor));
			}
		}

		private static void WriteHeight()
		{
			//HEIGHT
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Height") + Convert.ToString(mobjControl.Height * 20));
		}

		private static void WriteHelpContextID()
		{
			//HELPCONTEXTID
			if (mobjControl.HelpContextID != 0) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "HelpContextID") + Convert.ToString(mobjControl.HelpContextID));
			}
		}

		private static void WriteHideSelection()
		{
			//HIDESELECTION
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.HideSelection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.HideSelection == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "HideSelection") + "0  'False");
			}
		}

		private static void WriteIntegralHeight()
		{
			//INTEGRALHEIGHT
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.IntegralHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.IntegralHeight == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "IntegralHeight") + "0  'False");
			}
		}

		private static void WriteLargeChange()
		{
			//LARGECHANGE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.LargeChange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.LargeChange != 1) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.LargeChange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "LargeChange") + Convert.ToString(mobjControl.LargeChange));
			}
		}

		private static void WriteLeft()
		{
			int lngOffset = 0;
			if (mintContainer == 1)
				lngOffset = 30;
			//LEFT
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Left") + Convert.ToString(mobjControl.Left * 20 + lngOffset));
		}

		private static void WriteLocked()
		{
			//LOCKED
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Locked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Locked == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Locked") + "-1  'True");
			}
		}

		private static void WriteMaxScrollBar()
		{
			//MAX
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Max != 32767) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Max") + Convert.ToString(mobjControl.Max));
			}
		}

		private static void WriteMaxLength()
		{
			//MAXLENGTH
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MaxLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.MaxLength != 0) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MaxLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "MaxLength") + Convert.ToString(mobjControl.MaxLength));
			}
		}

		private static void WriteMinScrollBar()
		{
			//MIN
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Min != 0) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Min") + Convert.ToString(mobjControl.Min));
			}
		}

		private static void WriteMousePointer()
		{
			//MOUSEPOINTER
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			//UPGRADE_WARNING: modVfcMain property mobjControl.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			if (mobjControl.MousePointer != Microsoft.Vbe.Interop.Forms.fmMousePointer.fmMousePointerDefault) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MousePointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "MousePointer") + Convert.ToString(mobjControl.MousePointer));
			}
		}

		private static void WriteMultiLine()
		{
			//MULTILINE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MultiLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.MultiLine == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "MultiLine") + "-1  'True");
			}
		}

		private static void WriteMultiSelect()
		{
			//MULTISELECT
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MultiSelect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.MultiSelect != Microsoft.Vbe.Interop.Forms.fmMultiSelect.fmMultiSelectSingle) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MultiSelect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if (mobjControl.MultiSelect == Microsoft.Vbe.Interop.Forms.fmMultiSelect.fmMultiSelectMulti) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "MultiSelect") + "1  'Simple");
					//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.MultiSelect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				} else if (mobjControl.MultiSelect == Microsoft.Vbe.Interop.Forms.fmMultiSelect.fmMultiSelectExtended) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "MultiSelect") + "2  'Extended");
				}
			}
		}

		private static void WritePasswordChar()
		{
			//PASSWORDCHAR
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.PasswordChar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (!string.IsNullOrEmpty(mobjControl.PasswordChar)) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.PasswordChar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "PasswordChar") + FormatString(ref mobjControl.PasswordChar));
			}
		}

		private static void WriteScrollBars()
		{
			//SCROLLBARS
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ScrollBars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.ScrollBars != Microsoft.Vbe.Interop.Forms.fmScrollBars.fmScrollBarsNone) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ScrollBars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				if (mobjControl.ScrollBars == Microsoft.Vbe.Interop.Forms.fmScrollBars.fmScrollBarsHorizontal) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "ScrollBars") + "1  'Horizontal");
					//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ScrollBars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				} else if (mobjControl.ScrollBars == Microsoft.Vbe.Interop.Forms.fmScrollBars.fmScrollBarsVertical) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "ScrollBars") + "2  'Vertical");
					//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ScrollBars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				} else if (mobjControl.ScrollBars == Microsoft.Vbe.Interop.Forms.fmScrollBars.fmScrollBarsBoth) {
					FileSystem.PrintLine(mintFile, FormatProperty(ref "ScrollBars") + "3  'Both");
				}
			}
		}

		private static void WriteSmallChange()
		{
			//SMALLCHANGE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.SmallChange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.SmallChange != 1) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.SmallChange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "SmallChange") + Convert.ToString(mobjControl.SmallChange));
			}
		}

		private static void WriteStretch()
		{
			//STRETCH
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.PictureSizeMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.PictureSizeMode != Microsoft.Vbe.Interop.Forms.fmPictureSizeMode.fmPictureSizeModeClip) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Stretch") + "-1  'True");
			}
		}

		private static void WriteStyleComboBox()
		{
			//STYLE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Style != Microsoft.Vbe.Interop.Forms.fmStyle.fmStyleDropDownCombo) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Style") + "2  'Dropdown List");
			}
		}

		private static void WriteStyleCommandButton()
		{
			//STYLE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Picture != 0) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Style") + "1  'Graphical");
			}
		}

		private static void WriteStyleListBox()
		{
			//STYLE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.ListStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.ListStyle != Microsoft.Vbe.Interop.Forms.fmListStyle.fmListStylePlain) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Style") + "1  'Checkbox");
			}
		}

		private static void WriteStyleToggleButton()
		{
			//STYLE
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Style") + "1  'Graphical");
		}

		private static void WriteTabIndex()
		{
			//TABINDEX
			FileSystem.PrintLine(mintFile, FormatProperty(ref "TabIndex") + Convert.ToString(mobjControl.TabIndex));
		}

		private static void WriteTabStop()
		{
			//TABSTOP
			if (mobjControl.TabStop == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "TabStop") + "0  'False");
			}
		}

		private static void WriteTag()
		{
			//TAG
			if (!string.IsNullOrEmpty(mobjControl.Tag)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Tag") + FormatString(ref (mobjControl.Tag)));
			}
		}

		private static void WriteText()
		{
			//TEXT
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (!string.IsNullOrEmpty(mobjControl.Text)) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Text") + Strings.Left(FormatString(ref mobjControl.Text), 2047));
			}
		}

		private static void WriteToolTipText()
		{
			//TOOLTIPTEXT
			if (!string.IsNullOrEmpty(mobjControl.ControlTipText)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "ToolTipText") + FormatString(ref (mobjControl.ControlTipText)));
			}
		}

		private static void WriteTop()
		{
			int lngOffset = 0;
			if (mintContainer == 1)
				lngOffset = 120;
			//TOP
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Top") + Convert.ToString(mobjControl.Top * 20 + lngOffset));
		}

		private static void WriteValueCheckBox()
		{
			//VALUE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Value == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Value") + "1  'Checked");
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			} else if (Information.IsDBNull(mobjControl.Value)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Value") + "2  'Grayed");
			}
		}

		private static void WriteValueOptionButton()
		{
			//VALUE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Value == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Value") + "-1  'True");
			}
		}

		private static void WriteValueScrollBar()
		{
			//VALUE
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Value != 0) {
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Value") + Convert.ToString(mobjControl.Value));
			}
		}

		private static void WriteVisible()
		{
			//VISIBLE
			if (mobjControl.Visible == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Visible") + "0  'False");
			}
		}

		private static void WriteWidth()
		{
			//WIDTH
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Width") + Convert.ToString(mobjControl.Width * 20));
		}

		private static void WriteFormHeader()
		{
			FileSystem.PrintLine(mintFile, "VERSION 5.00");
			FileSystem.PrintLine(mintFile, "Begin VB.Form " + mobjMSForm.Name);
		}

		private static void WriteFormProperties()
		{
			Microsoft.Vbe.Interop.Forms.UserForm objUserForm = default(Microsoft.Vbe.Interop.Forms.UserForm);

			objUserForm = mobjMSForm.Designer;
			mintIndent = 1;

			//BACKCOLOR
			//UPGRADE_WARNING: Couldn't resolve default property of object objUserForm.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (Convert.ToDouble(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(objUserForm.BackColor))) != 0x8000000f) {
				//UPGRADE_WARNING: Couldn't resolve default property of object objUserForm.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "BackColor") + FormatHex(ref Convert.ToInt32(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(objUserForm.BackColor)))));
			}

			//BORDERSTYLE - Set to fixed single
			//This property although not included in VBA causes the VB form to act
			//like a VBA form, make changes to this property after importing
			FileSystem.PrintLine(mintFile, FormatProperty(ref "BorderStyle") + "1  'Fixed Single");

			//CAPTION
			if (!string.IsNullOrEmpty(objUserForm.Caption)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Caption") + FormatString(ref objUserForm.Caption));
			}

			//CLIENTHEIGHT
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ClientHeight") + Convert.ToString(objUserForm.InsideHeight * 20));

			//-CLIENTLEFT
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ClientLeft") + Convert.ToString(mobjMSForm.Properties.Item("Left").Value * 20 + 45));

			//-CLIENTTOP
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ClientTop") + Convert.ToString(mobjMSForm.Properties.Item("Top").Value * 20 + 330));

			//CLIENTWIDTH
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ClientWidth") + Convert.ToString(objUserForm.InsideWidth * 20));

			//ENABLED
			if (objUserForm.Enabled == false) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Enabled") + "0");
			}

			//FONT
			WriteFont(ref objUserForm.Font.Name);

			//FORECOLOR
			//UPGRADE_WARNING: Couldn't resolve default property of object objUserForm.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (Convert.ToDouble(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(objUserForm.ForeColor))) != 0x80000012) {
				//UPGRADE_WARNING: Couldn't resolve default property of object objUserForm.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileSystem.PrintLine(mintFile, FormatProperty(ref "ForeColor") + FormatHex(ref Convert.ToInt32(System.Drawing.ColorTranslator.FromOle(System.Convert.ToInt32(objUserForm.ForeColor)))));
			}

			//-HELPCONTEXTID
			if (mobjMSForm.Properties.Item("HelpContextID").Value != 0) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "HelpContextID") + Convert.ToString(mobjMSForm.Properties.Item("HelpContextID").Value));
			}

			//MAXBUTTON
			//This property although not included in VBA causes the VB form to act
			//like a VBA form, make changes to this property after importing
			FileSystem.PrintLine(mintFile, FormatProperty(ref "MaxButton") + "0   'False");

			//MINBUTTON
			//This property although not included in VBA causes the VB form to act
			//like a VBA form, make changes to this property after importing
			FileSystem.PrintLine(mintFile, FormatProperty(ref "MinButton") + "0   'False");

			//MOUSEPOINTER
			//UPGRADE_WARNING: modVfcMain property objUserForm.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			if (!objUserForm.MousePointer.@equals(Microsoft.Vbe.Interop.Forms.fmMousePointer.fmMousePointerDefault)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "MousePointer") + Convert.ToString(objUserForm.MousePointer));
			}

			//SCALEHEIGHT
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ScaleHeight") + Convert.ToString(objUserForm.InsideHeight * 20));

			//SCALEWIDTH
			FileSystem.PrintLine(mintFile, FormatProperty(ref "ScaleWidth") + Convert.ToString(objUserForm.InsideWidth * 20));

			//-STARTUPPOSITION
			switch (mobjMSForm.Properties.Item("StartUpPosition").Value) {
				case 0:
					FileSystem.PrintLine(mintFile, FormatProperty(ref "StartUpPosition") + "0  'Manual");
					break;
				case 1:
					FileSystem.PrintLine(mintFile, FormatProperty(ref "StartUpPosition") + "1  'CenterOwner");
					break;
				case 2:
					FileSystem.PrintLine(mintFile, FormatProperty(ref "StartUpPosition") + "2  'CenterScreen");
					break;
			}

			//-TAG
			if (!string.IsNullOrEmpty(mobjMSForm.Properties.Item("Tag").Value)) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "Tag") + FormatString(ref (mobjMSForm.Properties.Item("Tag").Value)));
			}

			//-WHATSTHISBUTTON
			if (mobjMSForm.Properties.Item("WhatsThisButton").Value == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "WhatsThisButton") + "-1  'True");
			}

			//-WHATSTHISHELP
			if (mobjMSForm.Properties.Item("WhatsThisHelp").Value == true) {
				FileSystem.PrintLine(mintFile, FormatProperty(ref "WhatsThisHelp") + "-1  'True");
			}

			//UPGRADE_NOTE: Object objUserForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objUserForm = null;
		}

		private static void WriteFormFooter()
		{
			FileSystem.PrintLine(mintFile, "End");
			FileSystem.PrintLine(mintFile, "Attribute VB_Name = \"" + mobjMSForm.Name + "\"");
			//TODO: figure out if these are the right settings to use
			FileSystem.PrintLine(mintFile, "Attribute VB_GlobalNameSpace = False");
			FileSystem.PrintLine(mintFile, "Attribute VB_Creatable = False");
			FileSystem.PrintLine(mintFile, "Attribute VB_PredeclaredId = True");
			FileSystem.PrintLine(mintFile, "Attribute VB_Exposed = False");
		}

		private static void WriteFormCode()
		{
			int lngLine = 0;
			if (mblnIncludeCode) {
				for (lngLine = 1; lngLine <= mobjMSForm.CodeModule.CountOfLines; lngLine++) {
					FileSystem.PrintLine(mintFile, mobjMSForm.CodeModule.Lines(lngLine, 1));
				}
			}
		}

		private static void WriteFont(ref System.Drawing.Font objFont)
		{
			string strProperty = null;
			FileSystem.PrintLine(mintFile, Indent() + "BeginProperty Font");
			mintIndent = mintIndent + 1;
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Name") + FormatString(ref (objFont.Name)));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Size") + Strings.Replace(Convert.ToString(objFont.SizeInPoints), ",", "."));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Charset") + Convert.ToString(objFont.GdiCharSet()));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Weight") + (objFont.Bold ? "700" : "400"));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Underline") + (objFont.Underline ? "-1  'True" : "0  'False"));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Italic") + (objFont.Italic ? "-1  'True" : "0  'False"));
			FileSystem.PrintLine(mintFile, FormatProperty(ref "Strikethrough") + (objFont.Strikeout ? "-1  'True" : "0  'False"));
			mintIndent = mintIndent - 1;
			FileSystem.PrintLine(mintFile, Indent() + "EndProperty");
		}

		private static bool ParentFontDifferent()
		{
			bool functionReturnValue = false;
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (mobjControl.Parent.Font.Name != mobjControl.Font.Name) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Size != mobjControl.Font.Size) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Charset != mobjControl.Font.Charset) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Weight != mobjControl.Font.Weight) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Underline != mobjControl.Font.Underline) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Italic != mobjControl.Font.Italic) {
				functionReturnValue = true;
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				//UPGRADE_WARNING: Couldn't resolve default property of object mobjControl.Parent.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			} else if (mobjControl.Parent.Font.Strikethrough != mobjControl.Font.Strikethrough) {
				functionReturnValue = true;
			}
			return functionReturnValue;
		}

		private static string FormatProperty(ref string strPropName)
		{
			string functionReturnValue = null;
			if (Strings.Len(strPropName) < 16) {
				functionReturnValue = Indent() + strPropName + Strings.Space(16 - Strings.Len(strPropName)) + "=   ";
			} else {
				functionReturnValue = Indent() + strPropName + " =   ";
			}
			return functionReturnValue;
		}

		private static string FormatString(ref string strValue)
		{
			string strChar = null;
			string strTemp = null;
			short intPos = 0;

			for (intPos = 1; intPos <= Strings.Len(strValue); intPos++) {
				strChar = Strings.Mid(strValue, intPos, 1);
				if (Strings.Asc(strChar) == 34) {
					strTemp = strTemp + Strings.Chr(34);
				}
				if (strChar == Constants.vbCr | strChar == Constants.vbLf) {
					strTemp = strTemp + "_";
				} else {
					strTemp = strTemp + strChar;
				}
			}

			return Strings.Chr(34) + strTemp + Strings.Chr(34);
		}

		private static string FormatHex(ref int lngValue)
		{
			return "&H" + Conversion.Hex(lngValue) + "&";
		}

		private static string Indent()
		{
			return Strings.Space(mintIndent * 3);
		}




//////////////////////////////////////////////////////

		public static object FindAcadVersionAndPlatformAndVertical()
		{
			//get the version by checking the title bar
			string[] titleBarSplited = null;
			titleBarSplited = Strings.Split(AutoCADAcadApplication_definst.Application.Caption, " ");
			short i = 0;
			mstrAcadVertical = "";
			for (i = 0; i <= Information.UBound(titleBarSplited); i++) {
				//first will find the vertical name,
				//then find the version number (year) and then exit for
				if ((Information.IsNumeric(titleBarSplited[i]))) {
					mstrAcadVersion = Convert.ToInt16(titleBarSplited[i]);
					break; // TODO: might not be correct. Was : Exit For
				} else if ((titleBarSplited[i] != "AutoCAD")) {
					//vertical names can have more than 1 word
					mstrAcadVertical = mstrAcadVertical + titleBarSplited[i];
				}
			}

			//get the if is 32 or 64 bit by checking the PLATFORM variable
			string platform = null;
			//UPGRADE_WARNING: Couldn't resolve default property of object Application.ActiveDocument.GetVariable(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			platform = AutoCADAcadApplication_definst.Application.ActiveDocument.GetVariable("PLATFORM");
			if ((Strings.InStr(1, platform, "x86") > 0)) {
				mstrAcadPlatform = 32;
			} else {
				mstrAcadPlatform = 64;
			}
			Debug.Print("AutoCAD " + mstrAcadVersion + " " + mstrAcadVertical + " " + mstrAcadPlatform + " bit");
		}

//Export all files in a VBA project and create a VB6 project to wrap them
		public static object ProcessProject()
		{
			Microsoft.Vbe.Interop.VBProject proj = default(Microsoft.Vbe.Interop.VBProject);
			Microsoft.Vbe.Interop.VBComponent comp = default(Microsoft.Vbe.Interop.VBComponent);
			string strFilePath = null;

			//Grab selected project and store its filepath to use while processing
			//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			proj = AutoCADAcadApplication_definst.Application.VBE.VBProjects(My.MyProject.Forms.frmVfcMain.SelectedProjectIndex);
			strFilePath = My.MyProject.Forms.frmVfcMain.SelectedProjectGitIgnoreDirectory;

			//We'll be creating a minimal VB6 project file to wrap together the exported modules.
			//strTxt contains the text for the project file
			string strTxt = null;

			strTxt = "Type=OleDll" + Constants.vbCrLf;
			//
			//    'If user wants us to add ObjectDBX Type Library
			//    If frmVfcMain.chkAddDbx = True Then
			//        'Explicitly add ObjectDBX Type Library as first TLB reference in project file. There are two reasons we do this:
			//        '1. Most VBA projects won't reference this (its not needed in VBA), but most VB6/.NET projects will.
			//        '     (And it doesn't matter if there are duplicates).
			//        '2. Adding it first ensures the Visual Studio conversion Wizard recognizes (for example) AcadLine as coming from the
			//        '     Autodesk.AutoCAD.Interop.Common namespace, and not Autodesk.AutoCAD.Interop namespace. (Less editing for you later).
			//        Call FindAcadVersionAndPlatformAndVertical
			//        If (mstrAcadPlatform = 32) Then
			//            Select Case mstrAcadVersion
			//                Case 2010 To 2012
			//                    strTxt = strTxt & "Reference=*\G{9F83C3E8-AAA3-4B0B-A256-F0DF98DA74BC}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb18enu.tlb#AXDBLib" & vbCrLf
			//                Case 2007 To 2009
			//                    strTxt = strTxt & "Reference=*\G{11A32D00-9E89-4C16-82CB-629DEBA56AE2}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb17enu.tlb#AXDBLib" & vbCrLf
			//                Case Else
			//                    MsgBox "Unfortunately this AutoCAD version is not supported", vbCritical
			//            End Select
			//        Else
			//            Select Case mstrAcadVersion
			//                Case 2010 To 2012
			//                    strTxt = strTxt & "Reference=*\G{FFC2A8DB-A497-4087-941C-C5B5462237EA}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb18enu.tlb#AXDBLib" & vbCrLf
			//                Case 2007 To 2009
			//                    strTxt = strTxt & "Reference=*\G{B789BF0E-B4A5-46B2-A8FE-D8CE0DA25E63}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb17enu.tlb#AXDBLib" & vbCrLf
			//                Case Else
			//                    MsgBox "Unfortunately this AutoCAD version is not supported", vbCritical
			//            End Select
			//        End If
			//    End If

			//add aditional references specific for verticals
			switch (mstrAcadVertical) {
				case "Civil3D":
					if ((Interaction.MsgBox("This VBA project is running on Civil 3D. Would you like to include aditional references? (AecXUIBase)", MsgBoxStyle.YesNo) == MsgBoxResult.Yes)) {
						switch (mstrAcadVersion) {
							case 2009:
								//only 32 bit
								strTxt = strTxt + "Reference=*\\G{E7BBE100-BF69-431B-9153-1FF0DEF8F613}#5.7#0#C:\\Program Files\\Common Files\\Autodesk Shared\\AecXUIBase57.tlb#AecXUIBase" + Constants.vbCrLf;
								break;
							case 2010:
								//only 32 bit
								strTxt = strTxt + "Reference=*\\G{E7BCE100-BF69-431B-9153-1FF0DEF8F613}#6.0#0#C:\\Program Files\\Common Files\\Autodesk Shared\\AecXUIBase60.tlb#AecXUIBase" + Constants.vbCrLf;
								break;
							case 2011:
								if ((mstrAcadPlatform == 32)) {
									//TODO
								//64 bits
								} else {
									//TODO
								}
								break;
							default:
								Interaction.MsgBox("Unfortunately this Civil 3D version is not supported.", MsgBoxStyle.Critical);
								break;
						}
					}
					break;
			}

			Microsoft.Vbe.Interop.Reference @ref = default(Microsoft.Vbe.Interop.Reference);

			foreach ( @ref in proj.References) {
				strTxt = strTxt + "Reference=*\\G" + @ref.GUID + "#" + @ref.Major + "." + @ref.Minor + "#0#" + @ref.FullPath + "#" + @ref.Name + Constants.vbCrLf;
			}


			//Iterate each component in project and export it with the right file extension
			//Note special processing for UserForms, which can't be migrated bby .NET Migration Wizard
			//  (have to convert them to VB6 Forms, which is most of the code in the redt of this project)
			foreach ( comp in proj.VBComponents) {
				if (comp.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm) {
					//Process a form
					ProcessForm(ref comp.Name, ref true);
					comp.Export((strFilePath + comp.Name + ".frm"));
					strTxt = strTxt + "Form=" + comp.Name + ".frmcmp" + Constants.vbCrLf;
				} else if (comp.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule) {
					//Process a Class module
					comp.Export((strFilePath + comp.Name + ".cls"));
					strTxt = strTxt + "Class=" + comp.Name + "; " + comp.Name + ".cls" + Constants.vbCrLf;
				} else if (comp.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule) {
					//Process a module
					comp.Export((strFilePath + comp.Name + ".bas"));
					strTxt = strTxt + "Module=" + comp.Name + "; " + comp.Name + ".bas" + Constants.vbCrLf;
				} else if (comp.Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document & comp.Name == "ThisDrawing") {
					//Process ThisDrawing module (which is exported as a class module)
					comp.Export((strFilePath + comp.Name + ".cls"));
					strTxt = strTxt + "Class=" + comp.Name + "; " + comp.Name + ".cls" + Constants.vbCrLf;
				}
			}

			//Save project file
			//This is a minimal VB6 project -
			//Remember we're not creating a project to use in VB6, we just want something we can
			//  give to the .NET Migration Wizard.

			FileSystem.FileOpen(42, strFilePath + proj.Name + ".vbp", OpenMode.Output);
			FileSystem.PrintLine(42, strTxt);
			FileSystem.FileClose(42);

			Debug.Print("Project export finished" + Constants.vbCrLf + "Project is in folder:" + Constants.vbCrLf + strFilePath);
		}


//Public Function PostProcessDotNetProject()
//
//    'If we've selected a project, then we take its directory as our default directory
//    If frmVfcMain.SelectedProjectFilename <> "" Then
//        frmVfcMain.ctrlFileDia.InitDir = frmVfcMain.SelectedProjectFilename
//    End If
//
//    'Ask user to select a file
//    On Error GoTo TheEnd
//    frmVfcMain.ctrlFileDia.DialogTitle = "Select vbproj file to post-process ..."
//    frmVfcMain.ctrlFileDia.DefaultExt = ".vbproj"
//    frmVfcMain.ctrlFileDia.Filter = ".vbproj"
//    frmVfcMain.ctrlFileDia.CancelError = True
//    frmVfcMain.ctrlFileDia.ShowOpen
//    On Error GoTo 0
//
//    'Open document .vbproj file
//    Dim doc As New DOMDocument
//    Dim node As IXMLDOMNode
//
//    If Dir(frmVfcMain.ctrlFileDia.fileName) = "" Then Exit Function
//
//    'Open XML vbproj file.
//    doc.Load frmVfcMain.ctrlFileDia.fileName
//    doc.async = False
//
//    'Set option to launch AutoCAD when we 'F5' debug (we add this to every configuration)
//    '(Note - Normally, VS adds this setting to your vbproj.user file, but it works if you modify the .vbproj file and delete any vbproj.user file that existed before you made the change.
//    '  You can also run this post-processing routine on the .vbproj.user file).
//    For Each node In doc.selectNodes("//PropertyGroup[@Condition]")
//        Dim newNode As IXMLDOMNode
//        Set newNode = node.appendChild(doc.createNode(NODE_ELEMENT, "StartAction", doc.firstChild.namespaceURI))
//        newNode.Text = "Program"
//        Set newNode = node.appendChild(doc.createNode(NODE_ELEMENT, "StartProgram", doc.firstChild.namespaceURI))
//        '*** Change pathname to match the installed location of AutoCAD 20XX on your machine ***
//        Dim verticalPath As String
//        Select Case mstrAcadVertical
//            Case "Civil3D"
//                verticalPath = "Civil 3D "
//        End Select
//        newNode.Text = "C:\Program Files\AutoCAD " & verticalPath & mstrAcadVersion & "\acad.exe"
//
//        'Try remove the x64 bit tags of projects created with Visual Basic Express
//        'http://msdn.microsoft.com/library/we1f72fb.aspx
//        Dim subNode As IXMLDOMNode
//        Dim subNodeConstant As IXMLDOMNode
//        For Each subNode In node.childNodes
//            If (subNode.nodeName = "PlatformTarget") Then
//                subNode.Text = "AnyCPU"
//            ElseIf (subNode.nodeName = "DefineConstants") Then
//                Set subNodeConstant = subNode
//            End If
//        Next
//        Call node.removeChild(subNodeConstant)
//    Next
//
//
//    'We're using VB Express 2008 and AutoCAD 2010, so we want to target Framework 3.5
//    '*** Comment this out if you want to target Framework 2.0 ***
//    Dim nodes As IXMLDOMNodeList
//    Dim newRefNode As IXMLDOMNode
//    Set nodes = doc.selectNodes("//PropertyGroup[ProjectType]")
//    Set node = nodes.Item(0)
//    If Not node Is Nothing Then
//        Set newRefNode = doc.createNode(NODE_ELEMENT, "TargetFrameworkVersion", doc.firstChild.namespaceURI)
//        newRefNode.Text = "v3.5"
//        node.appendChild newRefNode
//    End If
//
//
//    'Add references to acmgd.dll and acdbmgd.dll
//    '*** Edit the text we add below for different DLL versions ***
//    '(Easiest way to find text is to add references manually and then open up the vbproj file in notepad).
//    'We reference the one's from the ObjectARX SDK (because these are better) if we can find them,
//    'otherwise we reference the ones installed with AutoCAD (assuming default install location).
//
//    Dim strRefPath As String
//    strRefPath = frmVfcMain.txtARXSDKLocation
//    If Dir(frmVfcMain.txtARXSDKLocation, vbDirectory) = "" Then
//        'If textbox is empty, assume AutoCAD is installed in default location and use DLLs installed there. (Not ideal).
//        strRefPath = "C:\Program Files\AutoCAD " & verticalPath & mstrAcadVersion & "\"
//    End If
//
//    Set nodes = doc.selectNodes("//ItemGroup[Reference]")
//    Set node = nodes.Item(0)
//    If Not node Is Nothing Then
//
//        Set newRefNode = doc.createNode(NODE_ELEMENT, "Reference", doc.firstChild.namespaceURI)
//        Dim NewAttNode As IXMLDOMNode
//        Set NewAttNode = newRefNode.Attributes.setNamedItem(doc.createAttribute("Include"))
//        NewAttNode.Text = "acmgd, Version=18.0.0.0, Culture=neutral, processorArchitecture=x86"
//        Dim subRefNode As IXMLDOMNode
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "SpecificVersion", doc.firstChild.namespaceURI))
//        subRefNode.Text = "False"
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "HintPath", doc.firstChild.namespaceURI))
//        subRefNode.Text = strRefPath & "acmgd.dll"
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "Private", doc.firstChild.namespaceURI))
//        subRefNode.Text = "False"
//        node.appendChild newRefNode
//
//        Set newRefNode = doc.createNode(NODE_ELEMENT, "Reference", doc.firstChild.namespaceURI)
//        Set NewAttNode = newRefNode.Attributes.setNamedItem(doc.createAttribute("Include"))
//        NewAttNode.Text = "acdbmgd, Version=18.0.0.0, Culture=neutral, processorArchitecture=x86"
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "SpecificVersion", doc.firstChild.namespaceURI))
//        subRefNode.Text = "False"
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "HintPath", doc.firstChild.namespaceURI))
//        subRefNode.Text = strRefPath & "acdbmgd.dll"
//        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "Private", doc.firstChild.namespaceURI))
//        subRefNode.Text = "False"
//        node.appendChild newRefNode
//
//End If
//
//    'Save new .vbproj file
//    doc.Save frmVfcMain.ctrlFileDia.fileName
//    Set doc = Nothing
//
//    MsgBox "Finished processing .vbproj file", vbInformation, "VBA Converter"
//
//    Exit Function
//
//TheEnd:
//Err.Clear
//End Function
	}
}
