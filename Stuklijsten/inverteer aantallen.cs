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
			ScriptRecordset rsItemI = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItemI.MoveFirst();
			rsItemI.UseDataChanges = true;

			decimal aantal = Convert.ToDecimal(rsItemI.Fields["QUANTITY"].Value)*-1;

			rsItemI.Fields["PAINTAREA"].Value = 0.00;
            rsItemI.Fields["QUANTITY"].Value = aantal;

			rsItemI.Update();
		}
	}
}
