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
	static class UpgradeSupport
	{
		static internal AutoCAD.AcadApplication AutoCADAcadApplication_definst = new AutoCAD.AcadApplication();
		static internal TLI.TLIApplication TLITLIApplication_definst = new TLI.TLIApplication();
	}
}
