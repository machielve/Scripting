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

		/*
		Mutatie correctie
		

		*/

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
		{
			MessageBox.Show("Geen regels geselecteerd");
			return;
		}

		foreach (IRecord record in records)
		{
			ScriptRecordset rsStockIn = this.GetRecordset("R_STOCKIN", "", "PK_R_STOCKIN = " + (int)record.GetPrimaryKeyValue(), "");
			rsStockIn.MoveFirst();
			rsStockIn.UseDataChanges = true;

			string Naam = rsStockIn.Fields["DESCRIPTION"].Value.ToString();
			double Aantal = Convert.ToDouble(rsStockIn.Fields["QUANTITY"].Value.ToString());
			int regelId = (int)record.GetPrimaryKeyValue();

			string UitNaam = "Correctie - " + Naam;


			if (Naam.Substring(0, 15) != "Voorraadmutatie")
			{
				MessageBox.Show("Geen voorraadmutatie geselecteerd");
				return;
			}



			else
			{
				ScriptRecordset rsStockOut = this.GetRecordset("R_STOCKOUT", "", "PK_R_STOCKOUT = -1", "");
				rsStockOut.UseDataChanges = true;
				rsStockOut.AddNew();

				rsStockOut.Fields["FK_ITEM"].Value = rsStockIn.Fields["FK_ITEM"].Value;
				rsStockOut.Fields["QUANTITY"].Value = rsStockIn.Fields["QUANTITY"].Value;
				rsStockOut.Fields["DESCRIPTION"].Value = UitNaam;
				rsStockOut.Fields["FK_STOCKIN"].Value = regelId;

				rsStockIn.Fields["MEMO"].Value = "Voorraadmutatie verwijderd";

				rsStockOut.Update();
				rsStockIn.Update();

			}



		}





	}
}