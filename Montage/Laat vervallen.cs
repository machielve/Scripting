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
			het laten vervallen van de geselecteerde montage instructies

		*/


		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
		{
			MessageBox.Show("Geen instructies geselecteerd");
			return;
		}


		foreach (IRecord record in records)
		{
			ScriptRecordset rsInst = this.GetRecordset("U_MONTAGE_STRUCTIES", "", "PK_U_MONTAGE_STRUCTIES = " + (int)record.GetPrimaryKeyValue(), "");
			rsInst.MoveFirst();

			rsInst.Fields["STATUS_MONTAGE_INSTRUCTIE"].Value = 2;
			
			rsInst.Update();

		}

		MessageBox.Show("Geselecteerde instructies zijn vervallen");


	}
}