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
		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsISUP = this.GetRecordset("U_PACKLIST", "", "PK_U_PACKLIST = " + (int)record.GetPrimaryKeyValue(), "");
			rsISUP.MoveFirst();

			rsISUP.Fields["GEREED"].Value = false;

			rsISUP.Update();
		}
	}
}