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

		string pad = @"W:\Almacon Tekeningen\ALP2021\";
		string t2 = Tekening1 + ".pdf";

		string filename = pad + groep + @"\" + t2;

		if (File.Exists(filename))

		{
			System.Diagnostics.Process.Start(filename);

		}
		else MessageBox.Show(@"Bestand niet te vinden");

	}
}