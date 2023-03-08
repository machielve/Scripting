using ADODB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Xml;
using Ridder.Common.ADO;
using Ridder.Common.Choices;
using Ridder.Common.Search;
using System.Linq;
using Ridder.Recordset.Extensions;
using System.Windows.Forms;
using System.Data;
using Ridder.Common.Script;
using System.IO;

public class RidderScript : CommandScript
{
	public void Execute()
	{
		string oldLocation = @"C:\Program Files\Bricsys\BricsCAD V21 en_US\Support\";
		string BULocation = @"W:\Machiel\BricsCAD\Lisp\";
		string newLocation = @"C:\Users\machiel.vanemden\AppData\Roaming\Bricsys\BricsCAD\V23x64\en_US\Support\";

		string fileName = @"on_start.lsp";
		string backup1 = @"on_start-backup21.lsp";
		string backup2 = @"on_start-backup23.lsp";


		string oldFile21 = oldLocation + fileName;
		string oldFile23 = newLocation + fileName;
		string BUFile21 = BULocation + backup1;
		string BUFile23 = BULocation + backup2;
		string newFile = newLocation + fileName;

		if (File.Exists(oldFile23) )// && File.Exists(oldFile21))
		{
			
			//File.Copy(oldFile21, BUFile21);
			//File.Move(oldFile23, BUFile23);
			//File.Copy(oldFile21, newFile);

			
			MessageBox.Show("Bestand is er");
			
		}
		else MessageBox.Show("Is er niet");
	}
}
