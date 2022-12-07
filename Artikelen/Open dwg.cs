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
using System.ComponentModel;
using Ridder.Common.Script;
using System.IO;


public class RidderScript : CommandScript
{
	public void Execute()
	{
		string Tekening1 = this.FormDataAwareFunctions.CurrentRecord.GetCurrentRecordValue("DRAWINGNUMBER").ToString();
		string groep = Tekening1.Substring(0, 5);
		string t2 = Tekening1 + ".dwg";

		string pad = @"W:\Almacon Tekeningen\ALP2005\";
		string pad3 = @"W:\Almacon Tekeningen\ALP2021\";

		string filename = pad + groep +@"\"+ t2;
		string filename3 = pad3 + groep +@"\"+ t2;

		if (File.Exists(filename3))
		{
			System.Diagnostics.Process.Start(filename3);
			MessageBox.Show(@"Bestand geopend vanaf W:\ALP2021");
		}
		
		
		else if (File.Exists(filename))
		{
			MessageBox.Show(@"Bestand niet te vinden in W:\ALP2021");
			System.Diagnostics.Process.Start(filename);
			MessageBox.Show(@"Bestand geopend vanaf W:\ALP2005");
		}
		
		else MessageBox.Show(@"Bestand niet te vinden in W:\ALP2021 of W:\ALP2005");
		
	}
}