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
		var BonRID = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();

		ScriptRecordset rsBon = this.GetRecordset("R_JOBORDERDETAILITEM", "", string.Format("PK_R_JOBORDERDETAILITEM = '{0}'", BonRID.ToString()), "");
		rsBon.MoveFirst();

		string BonID = rsBon.Fields["FK_JOBORDER"].Value.ToString();

		ScriptRecordset Bonnen = this.GetRecordset("R_JOBORDERDETAILITEM", "FK_ITEM", string.Format("FK_JOBORDER = '{0}'", BonID), "PK_R_JOBORDERDETAILITEM");
		ScriptRecordset Bonnen1 = this.GetRecordset("R_JOBORDERDETAILITEM", "FK_ITEM", string.Format("FK_JOBORDER = '{0}'", BonID), "PK_R_JOBORDERDETAILITEM");
		int aantal = Bonnen.RecordCount;
		int aantal1 = Bonnen1.RecordCount;
		MessageBox.Show(aantal.ToString()+" / "+aantal1.ToString());
		
		Bonnen.MoveFirst();
		
		int i = 0;
		int j = 0;

		while (i < aantal)
		{
			string nummer = Bonnen.Fields["FK_ITEM"].Value.ToString();
			Bonnen1.MoveFirst();
			Bonnen1.MoveNext();

			while (j < aantal)
			{
				string nummer1 = Bonnen1.Fields["FK_ITEM"].Value.ToString();
				if (nummer == nummer1)
				{
					MessageBox.Show(nummer1);
				}
				Bonnen1.MoveNext();

				j++;

			}


			Bonnen.MoveNext();
			i++;
		}

		
		

	}
}