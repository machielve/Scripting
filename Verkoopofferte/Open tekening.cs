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

public class RidderScript : CommandScript
{
	public void Execute()
	{
		string Tekening = this.FormDataAwareFunctions.CurrentRecord.GetCurrentRecordValue("DRAWINGNUMBER").ToString();

		
		string groep = Tekening.Substring(0, 3);
		string map = groep + "00-" + groep + @"99\";
		string pad = @"W:\Almacon Tekeningen\";
		string t2 = Tekening + ".dwg";

		string filename = pad + map + t2;
		
		System.Diagnostics.Process.Start(filename);	

	}
}