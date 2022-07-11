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
    /*
	
	Voorraad reservering verwijderen, het  programma om een reservering van een voorraadhoudende bonregel artikel te verwijderen
	Uit te voeren vanuit een bonregel artikel welke niet verder mag zijn dan gereserveerd
	Geschreven door: Machiel R. van Emden mei-2022

	*/

	public void Execute()
	{

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsBonregel = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsBonregel.MoveFirst();
			//		rsBonregel.UseDataChanges = true;

			string ItemCode = rsBonregel.Fields["FK_ITEM"].Value.ToString();
			string omschrijving = rsBonregel.Fields["DESCRIPTION"].Value.ToString();
			double aantal = Convert.ToDouble(rsBonregel.Fields["QUANTITY"].Value.ToString());
			double lengte = Convert.ToDouble(rsBonregel.Fields["LENGTH"].Value.ToString());
			int bon = Convert.ToInt32(rsBonregel.Fields["FK_JOBORDER"].Value.ToString());

			rsBonregel.Delete();

			ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
			rsJoborderItem.UseDataChanges = true;
			rsJoborderItem.AddNew();

			rsJoborderItem.Fields["FK_JOBORDER"].Value = bon;
			rsJoborderItem.Fields["FK_ITEM"].Value = Convert.ToInt32(ItemCode);
			rsJoborderItem.Fields["DESCRIPTION"].Value = omschrijving;
			rsJoborderItem.Fields["QUANTITY"].Value = aantal;
			rsJoborderItem.Fields["LENGTH"].Value = lengte;

			rsJoborderItem.Update();

		}

		this.FormDataAwareFunctions.Refres();

	}

	// M.R.v.E - 2022
	
}