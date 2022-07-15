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

		ScriptRecordset rsBon = this.GetRecordset("R_JOBORDER", "", string.Format("PK_R_JOBORDER = '{0}'", BonRID.ToString()), "");
		rsBon.MoveFirst();

		string BonID = rsBon.Fields["PK_R_JOBORDER"].Value.ToString();

		ScriptRecordset Bonnen = this.GetRecordset("R_JOBORDERDETAILITEM", "", string.Format("FK_JOBORDER = '{0}'", BonID), "PK_R_JOBORDERDETAILITEM");
		ScriptRecordset Bonnen1 = this.GetRecordset("R_JOBORDERDETAILITEM", "", string.Format("FK_JOBORDER = '{0}'", BonID), "PK_R_JOBORDERDETAILITEM");
		int aantal = Bonnen.RecordCount;
		int aantal1 = Bonnen1.RecordCount;
		//MessageBox.Show(aantal.ToString()+" / "+aantal1.ToString());

		Bonnen.MoveFirst();

		int i = 0;
		int j = 0;

		while (Bonnen.EOF == false)
		{
			string nummer = Bonnen.Fields["FK_ITEM"].Value.ToString();
			string regel = Bonnen.Fields["PK_R_JOBORDERDETAILITEM"].Value.ToString();
			string lengte = Bonnen.Fields["LENGTH"].Value.ToString();
			string quant = Bonnen.Fields["QUANTITY"].Value.ToString();
			int quantity = Convert.ToInt32(quant);
			Bonnen1.MoveFirst();
			//Bonnen1.MoveNext();

			while (Bonnen1.EOF == false)
			{
				//MessageBox.Show("Test"+"-"+j.ToString());

				string nummer1 = Bonnen1.Fields["FK_ITEM"].Value.ToString();
				string regel1 = Bonnen1.Fields["PK_R_JOBORDERDETAILITEM"].Value.ToString();
				string lengte1 = Bonnen1.Fields["LENGTH"].Value.ToString();
				string quant1 = Bonnen1.Fields["QUANTITY"].Value.ToString();
				int quantity1 = Convert.ToInt32(quant1);

				if (nummer == nummer1 && regel != regel1 && lengte == lengte1)
				{
					int extra = quantity + quantity1;
					Bonnen.Fields["QUANTITY"].Value = extra;
					Bonnen.Update();
					Bonnen1.Delete();
					Bonnen1.Update();

					//MessageBox.Show("Gelijk - " + extra.ToString());
				}
				Bonnen.Update();
				Bonnen1.MoveNext();

				j++;

			}

			MessageBox.Show("Loop -" + i.ToString());

			Bonnen.MoveNext();
			i++;
			j = 0;
		}

	}
}