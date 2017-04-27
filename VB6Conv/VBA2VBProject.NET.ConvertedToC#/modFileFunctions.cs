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
	static class modFileFunctions
	{
		private const int BIF_RETURNONLYFSDIRS = 0x1;
		private const int BIF_DONTGOBELOWDOMAIN = 0x2;
		private const int BIF_RETURNFSANCESTORS = 0x8;
		private const int BIF_BROWSEFORCOMPUTER = 0x1000;
		private const int BIF_BROWSEFORPRINTER = 0x2000;
		private const int BIF_BROWSEINCLUDEFILES = 0x4000;
		private const int MAX_PATH = 260;

		public static string BrowseFolder(ref string Caption = "", ref string InitialFolder = "")
		{

			Shell32.Shell SH = null;
			Shell32.Folder F = null;

			SH = new Shell32.Shell();
			F = SH.BrowseForFolder(0, Caption, BIF_RETURNONLYFSDIRS, InitialFolder);
			if ((F != null)) {
				return F.Items().Item().Path;
			}

		}

//We store our output files in a subdirectory of the original project folder
//  so added a subfolder to the returned path.
		public static string GetFileDirectory(ref string strFileName)
		{
			string tmpString = "";
			short intPos = 0;
			for (intPos = Strings.Len(strFileName); intPos >= 1; intPos += -1) {
				if (Strings.Mid(strFileName, intPos, 1) == "\\") {
					tmpString = Strings.Left(strFileName, intPos);
					break; // TODO: might not be correct. Was : Exit For
				}
			}
			tmpString = tmpString + "VB6Conv\\";
			//Dim fso As New FileSystemObject
			//UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			if (string.IsNullOrEmpty(FileSystem.Dir(tmpString, FileAttribute.Directory)))
				FileSystem.MkDir(GetFileDirectory());
			return tmpString;

		}

		public static string GetFileNameFromPath(string strFilePath)
		{
			return Strings.Mid(strFilePath, Strings.InStrRev(strFilePath, "\\") + 1);
		}

		public static string StripExtensionFromFileName(string strFileName)
		{
			return Strings.Left(strFileName, Strings.InStrRev(strFileName, ".") - 1);

		}
	}
}
