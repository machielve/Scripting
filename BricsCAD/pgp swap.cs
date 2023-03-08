using ADODB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Xml;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;

public class puuuh
{
	public void Execute()
	{
		string oldLocation = @"C:\Users\machiel.vanemden\AppData\Roaming\Bricsys\BricsCAD\V21x64\en_US\Support\";
		string BULocation = @"W:\Machiel\BricsCAD\Herbert";
		string newLocation = @"C:\Users\machiel.vanemden\AppData\Roaming\Bricsys\BricsCAD\V23x64\en_US\Support\";

		string fileName = @"default.pgp";
		string backup1 = @"default-backup21.pgp";
		string backup2 = @"default-backup23.pgp";
	

		string oldFile21 = oldLocation + fileName;
		string oldFile23 = newLocation + fileName;
		string BUFile21 = BULocation + backup1;
		string BUFile23 = BULocation + backup2;
		string newFile = newLocation + fileName;

		if (File.Exists(oldFile21) && File.Exists(oldFile23))
		{
			File.Copy(oldFile21, BUFile21);
			File.Move(oldFile23, BUFile23);			
			File.Copy(oldFile21, newFile);

            Messagebox.show("Bestanden geregeld");

		}
        else MessageBox.show("Bron bestand bestaat niet");

	}
}