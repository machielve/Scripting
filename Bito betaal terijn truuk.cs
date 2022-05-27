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
using Ridder.Client.SDK;


public class RidderScript : CommandScript
{
	/*
	
	Bito truuk , het  programma om de betaaltermijnen te berekenen met de afspraak van een aanbetaling van 10.000
	Uit te voeren vanuit een order met bestaande open betaaltermijnen
	Geschreven door: Machiel R. van Emden mei-2022

	*/


	public void Execute()

	{

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;
		
		decimal totaal = 0;

		decimal Termijn2 = 0;

		int check1 = 0;
		int check2 = 0;
		int check3 = 0;
		
		

		foreach (IRecord record in records)
		{
			ScriptRecordset rsBetaalT = this.GetRecordset("R_SALESINSTALLMENT", "", "PK_R_SALESINSTALLMENT = " + (int)record.GetPrimaryKeyValue(), "");
			rsBetaalT.MoveFirst();
			rsBetaalT.UseDataChanges = true;

			string termijn = rsBetaalT.Fields["FK_INVOICESCHEDULEDETAIL"].Value.ToString();
			
			decimal prijs = Convert.ToDecimal(rsBetaalT.Fields["NETAMOUNT"].Value.ToString());

			if (termijn == "18")
			{
				check1 += 1;
			}
			if (termijn == "19")
			{
				Termijn2 += prijs;
				check2 += 1;
			}
			if (termijn == "20")
			{
				check3 += 1;
			}
		
			totaal += prijs;
				

			rsBetaalT.Update();

		}

		if (check1 == 0)
		{

			MessageBox.Show("Termijn 1 niet geselecteerd"); 
			return;
		
		}
		if (check2 == 0)
		{
			MessageBox.Show("Termijn 2 niet geselecteerd");
			return;

		}
		if (check3 == 0)
		{
			MessageBox.Show("Termijn 3 niet geselecteerd");
			return;

		}
		

		foreach (IRecord record in records)
		{
			ScriptRecordset rsBetaalT = this.GetRecordset("R_SALESINSTALLMENT", "", "PK_R_SALESINSTALLMENT = " + (int)record.GetPrimaryKeyValue(), "");
			rsBetaalT.MoveFirst();
			rsBetaalT.UseDataChanges = true;

			string termijn = rsBetaalT.Fields["FK_INVOICESCHEDULEDETAIL"].Value.ToString();

			decimal prijs = Convert.ToDecimal(rsBetaalT.Fields["NETAMOUNT"].Value.ToString());

			if (termijn == "20")
			{
				rsBetaalT.Fields["FINALINSTALLMENT"].Value = false;
				rsBetaalT.Fields["NETAMOUNT"].Value = 0;

			}

			rsBetaalT.Update();

		}


		foreach (IRecord record in records)
		{
			ScriptRecordset rsBetaalT = this.GetRecordset("R_SALESINSTALLMENT", "", "PK_R_SALESINSTALLMENT = " + (int)record.GetPrimaryKeyValue(), "");
			rsBetaalT.MoveFirst();
			rsBetaalT.UseDataChanges = true;

			string termijn = rsBetaalT.Fields["FK_INVOICESCHEDULEDETAIL"].Value.ToString();

			decimal prijs = Convert.ToDecimal(rsBetaalT.Fields["NETAMOUNT"].Value.ToString());

			if (termijn == "18")
			{
				rsBetaalT.Fields["NETAMOUNT"].Value = 10000;

			}

			rsBetaalT.Update();

		}

		foreach (IRecord record in records)
		{
			ScriptRecordset rsBetaalT = this.GetRecordset("R_SALESINSTALLMENT", "", "PK_R_SALESINSTALLMENT = " + (int)record.GetPrimaryKeyValue(), "");
			rsBetaalT.MoveFirst();
			rsBetaalT.UseDataChanges = true;

			string termijn = rsBetaalT.Fields["FK_INVOICESCHEDULEDETAIL"].Value.ToString();

			decimal prijs = Convert.ToDecimal(rsBetaalT.Fields["NETAMOUNT"].Value.ToString());

			if (termijn == "20")
			{
				decimal rest = totaal - 10000 - Termijn2;
				rsBetaalT.Fields["NETAMOUNT"].Value = rest;
				rsBetaalT.Fields["FINALINSTALLMENT"].Value = true;

			}

			rsBetaalT.Update();

		}
		
		
		

	//	MessageBox.Show(totaal.ToString());

		this.FormDataAwareFunctions.Refres();

	}
}