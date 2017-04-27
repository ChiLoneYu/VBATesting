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
	static class modDevTools
	{

//Must have a comma at the end of these file extension lists for InStr to work correctly
		public const string VBAImportTypes = "FRM,BAS,CLS,";
		public const string VBAFileTypes = "FRMCMP,FRM,BAS,CLS,FRX,";
		public const string VBAMoveTypes = "FRMCMP,BAS,CLS,VBP,";

		static Microsoft.Vbe.Interop.VBProject project;


		public static void ImportSourceFiles(ref string sourcePath)
		{
			string file = null;
			string fileExt = null;
			//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			file = FileSystem.Dir(sourcePath);
			while ((!string.IsNullOrEmpty(file))) {
				fileExt = GetFileExtension(ref file, ref true);
				//MsgBox fileExt
				if (Strings.InStr(VBAImportTypes, fileExt + ",") > 0) {
					//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject.VBComponents.Import(sourcePath + file);
				}

				//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				file = FileSystem.Dir();
			}
		}

		public static void ExportSourceFiles(ref string destPath)
		{
			clsCRC32 CRC32 = new clsCRC32();
			Microsoft.Vbe.Interop.VBComponent comp = default(Microsoft.Vbe.Interop.VBComponent);
			string ignoreFilePath = null;
			string cleanFilePath = null;
			string ignorePath = null;
			Scripting.FileSystemObject fileSystemHandler = new Scripting.FileSystemObject();
			Scripting.File ignoreFile = null;
			Scripting.File rootFile = null;
			string frxName = null;
			string frmName = null;
			bool copyFRX = false;


			//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			project = AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject;


			ignorePath = destPath + "ignore\\";
			//Clean out the ignore folder
			//fileSystemHandler.DeleteFile (ignorePath & "*")


			//No longer needed since modVfcMain is handling this...
			//Loop through all of the components and export to the ignore folder
			//    For Each comp In project.VBComponents
			//        If Left(comp.Name, 5) <> "Sheet" And comp.Name <> "ThisWorkbook" And comp.Name <> "DevTools" And comp.Name <> "ThisDrawing" Then
			//            'cleanFilePath = destPath & comp.Name & ToFileExtension(comp.Type)
			//            ignoreFilePath = ignorePath & comp.Name & ToFileExtension(comp.Type)
			//            comp.Export ignoreFilePath
			//        End If
			//    Next


			//Loop through root directory and find files with no match in ignore folder
			foreach (Scripting.File rootFile_loopVariable in fileSystemHandler.GetFolder(destPath).Files) {
				rootFile = rootFile_loopVariable;

				//Check that this file is a type handled by VBA
				if (Strings.InStr(VBAFileTypes, GetFileExtension(ref rootFile.Name, ref true) + ",") > 0) {

					ignoreFilePath = ignorePath + rootFile.Name;

					//If file doesn't exist in the ignore folder, then go ahead and delete it... most likely removed from VBProject
					if (!fileSystemHandler.FileExists(ignoreFilePath)) {
						rootFile.Delete();
					}
				}

			}

			//Loop through all of the files in the ignore directory and perform a CRC compare.  Copy to root if CRC doesn't match
			//Make sure to ignore .FRX

			foreach (Scripting.File ignoreFile_loopVariable in fileSystemHandler.GetFolder(ignorePath).Files) {
				ignoreFile = ignoreFile_loopVariable;

				copyFRX = false;

				if (Strings.InStr(VBAMoveTypes, GetFileExtension(ref ignoreFile.Name, ref true) + ",") > 0) {
					//Check that file exists in root, otherwise it should be copied by default
					//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					if (!string.IsNullOrEmpty(FileSystem.Dir(destPath + ignoreFile.Name))) {
						//Move the file if the CRC32 doesn't match.. otherwise delete it
						if (Conversion.Hex(CRC32.CalcCRC32(ref ignorePath + ignoreFile.Name)) != Conversion.Hex(CRC32.CalcCRC32(ref destPath + ignoreFile.Name))) {
							ignoreFile.Copy(destPath + ignoreFile.Name, true);

							//Also copy the matching FRM and FRX files if a form comparison file is found with a non-matching CRC32
							if (GetFileExtension(ref ignoreFile.Name, ref true) == "FRMCMP") {
								copyFRX = true;

							}

						}

					} else {
						//Copy from ignore to root since file doesn't exist
						ignoreFile.Copy(destPath, true);

						//Also copy the match FRX file if copying a form
						if (GetFileExtension(ref ignoreFile.Name, ref true) == "FRMCMP") {

							copyFRX = true;

						}

					}

					//ignoreFile.Delete

				}

				if (copyFRX & GetFileExtension(ref ignoreFile.Name, ref true) == "FRMCMP") {
					frxName = Strings.Left(ignoreFile.Name, Strings.Len(ignoreFile.Name) - 6) + "frx";
					frmName = Strings.Left(ignoreFile.Name, Strings.Len(ignoreFile.Name) - 6) + "frm";

					fileSystemHandler.CopyFile(ignorePath + frxName, destPath + frxName, true);
					fileSystemHandler.CopyFile(ignorePath + frmName, destPath + frmName, true);

					//fileSystemHandler.DeleteFile (ignorePath & frxName)
					//fileSystemHandler.DeleteFile (ignorePath & frmName)

				}


			}

			if (My.MyProject.Forms.frmVfcMain.cbDeleteIgnoreFiles.CheckState) {
				fileSystemHandler.DeleteFile(destPath + "ignore\\*");
			}

		}

		public static void RemoveAllModules()
		{
			Microsoft.Vbe.Interop.VBProject project = default(Microsoft.Vbe.Interop.VBProject);
			//UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			project = AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject;

			Interaction.MsgBox("Removing Modules");
			Microsoft.Vbe.Interop.VBComponent comp = default(Microsoft.Vbe.Interop.VBComponent);
			foreach ( comp in project.VBComponents) {
				//And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
				if (!(comp.Name == "DevTools") & Strings.Left(comp.Name, 5) != "Sheet" & comp.Name != "ThisWorkbook") {
					project.VBComponents.Remove(comp);
				}
			}
		}

		private static string ToFileExtension(ref Microsoft.Vbe.Interop.vbext_ComponentType vbeComponentType)
		{
			string functionReturnValue = null;
			switch (vbeComponentType) {
				case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
					functionReturnValue = ".cls";
					break;
				case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
					functionReturnValue = ".bas";
					break;
				case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
					functionReturnValue = ".frm";
					break;
				case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
					break;
				case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
					break;
				default:
					functionReturnValue = Constants.vbNullString;
					break;
			}
			return functionReturnValue;

		}

		private static string GetFileExtension(ref string fileName, ref bool upperCase = false)
		{
			string functionReturnValue = null;

			if (upperCase) {
				functionReturnValue = Strings.UCase(Strings.Right(fileName, Strings.Len(fileName) - Strings.InStrRev(fileName, ".")));
			} else {
				functionReturnValue = Strings.Right(fileName, Strings.Len(fileName) - Strings.InStrRev(fileName, "."));
			}
			return functionReturnValue;

		}
	}
}
