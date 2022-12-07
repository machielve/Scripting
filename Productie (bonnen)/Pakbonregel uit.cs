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
			ScriptRecordset rsItem = this.GetRecordset("R_JOBORDERDETAILITEM", "DIRECTELEVERING", "PK_R_JOBORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			ScriptRecordset rsPBR = this.GetRecordset("U_PACKLISTDETAILITEM", "", "FK_BONREGELART = " + (int)record.GetPrimaryKeyValue(), "");
			rsPBR.MoveFirst();

			if (rsPBR.RecordCount!=0)
			{
				MessageBox.Show("Pakbonregel kan niet uit, regel word op pakbon(nen) gebuikt");
				continue;
			}
			


			rsItem.Fields["DIRECTELEVERING"].Value = false;

			rsItem.Update(null, null);


		}


	}
}