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
	internal class clsCRC32
	{

		private int[] CRCTable = new int[256];

		public int CalcCRC32(ref string FilePath)
		{
			byte[] ByteArray = null;
			int Limit = 0;
			int CRC = 0;
			int Temp1 = 0;
			int Temp2 = 0;
			int i = 0;
			short intFF = 0;

			intFF = FreeFile();
			FileSystem.FileOpen(intFF, FilePath, OpenMode.Binary, OpenAccess.Read);
			Limit = FileSystem.LOF(intFF);
			ByteArray = new byte[Limit];
			//UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileSystem.FileGet(intFF, ByteArray);
			FileSystem.FileClose(intFF);

			Limit = Limit - 1;
			CRC = -1;
			for (i = 0; i <= Limit; i++) {
				if (CRC < 0) {
					Temp1 = CRC & 0x7fffffff;
					Temp1 = Temp1 / 256;
					Temp1 = (Temp1 | 0x800000) & 0xffffff;
				} else {
					Temp1 = (CRC / 256) & 0xffffff;
				}
				Temp2 = ByteArray[i];
				// get the byte
				Temp2 = CRCTable[(CRC ^ Temp2) & 0xff];
				CRC = Temp1 ^ Temp2;
			}
			CRC = CRC ^ 0xffffffff;
			return CRC;
		}

//UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		private void Class_Initialize_Renamed()
		{
			short i = 0;
			short J = 0;
			int Limit = 0;
			int CRC = 0;
			int Temp1 = 0;
			Limit = 0xedb88320;
			for (i = 0; i <= 255; i++) {
				CRC = i;
				for (J = 8; J >= 1; J += -1) {
					if (CRC < 0) {
						Temp1 = CRC & 0x7fffffff;
						Temp1 = Temp1 / 2;
						Temp1 = Temp1 | 0x40000000;
					} else {
						Temp1 = CRC / 2;
					}
					if (CRC & 1) {
						CRC = Temp1 ^ Limit;
					} else {
						CRC = Temp1;
					}
				}
				CRCTable[i] = CRC;
			}
		}
		public clsCRC32() : base()
		{
			Class_Initialize_Renamed();
		}
	}
}
