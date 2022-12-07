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

		Bonnen.MoveFirst();


		while (Bonnen.EOF == false)
		{
			string nummer = Bonnen.Fields["FK_ITEM"].Value.ToString();
			string regel = Bonnen.Fields["PK_R_JOBORDERDETAILITEM"].Value.ToString();
			string lengte = Bonnen.Fields["LENGTH"].Value.ToString();
			string quant = Bonnen.Fields["QUANTITY"].Value.ToString();
			int quantity = Convert.ToInt32(quant);
			Bonnen1.MoveFirst();

			while (Bonnen1.EOF == false)
			{
				string nummer1 = Bonnen1.Fields["FK_ITEM"].Value.ToString();
				string regel1 = Bonnen1.Fields["PK_R_JOBORDERDETAILITEM"].Value.ToString();
				string lengte1 = Bonnen1.Fields["LENGTH"].Value.ToString();
				string quant1 = Bonnen1.Fields["QUANTITY"].Value.ToString();
				int quantity1 = Convert.ToInt32(quant1);

				if (	nummer == nummer1 && 
						regel != regel1 && 
						lengte == lengte1)
				{
					int extra = quantity + quantity1;
					Bonnen.Fields["QUANTITY"].Value = extra;
					Bonnen.Update();
					Bonnen1.Delete();
					Bonnen1.Update();
				}
				
				else 
				{
					//Bonnen1.Update();
					//Bonnen.Update();
					if (Bonnen1.EOF==false) Bonnen1.MoveNext();
				}

			}

			if (Bonnen.EOF==false) Bonnen.MoveNext();
		}

	}
}