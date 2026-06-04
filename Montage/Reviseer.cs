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
			het Reviseren van de geselecteerde montage instructies
		*/


		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
		{
			MessageBox.Show("Geen instructies geselecteerd");
			return;
		}


		foreach (IRecord record in records)
		{
			// ophalen oude data
			ScriptRecordset rsInst = this.GetRecordset("U_MONTAGE_STRUCTIES", "", "PK_U_MONTAGE_STRUCTIES = " + (int)record.GetPrimaryKeyValue(), "");
			rsInst.MoveFirst();
			
			string omschrijving = rsInst.Fields["OMSCHRIJVING"].Value.ToString();
			string instructie = rsInst.Fields["INSTRUCTIE_TEKST"].Value.ToString();
			string onderdelen = rsInst.Fields["ONDERDELENLIJST"].Value.ToString();;

			object imageA = rsInst.Fields["AFBEELDING_A"].Value;
			object imageB = rsInst.Fields["AFBEELDING_B"].Value;
			object imageC = rsInst.Fields["AFBEELDING_C"].Value;
			object imageD = rsInst.Fields["AFBEELDING_D"].Value;
			object imageE = rsInst.Fields["AFBEELDING_E"].Value;
			object imageF = rsInst.Fields["AFBEELDING_F"].Value;
			object imageG = rsInst.Fields["AFBEELDING_G"].Value;
			object imageH = rsInst.Fields["AFBEELDING_H"].Value;			
				
			string versie = rsInst.Fields["VERSIE"].Value.ToString();
			int editie = Convert.ToInt32(versie);
			int editie2 = editie + 1;
			string versie2 = editie2.ToString();


			//maken van nieuwe data
			ScriptRecordset rsInstNew = this.GetRecordset("U_MONTAGE_STRUCTIES", "", "PK_U_MONTAGE_STRUCTIES= -1", "");
			rsInstNew.UseDataChanges = true;
			rsInstNew.AddNew();

			rsInstNew.Fields["VERSIE"].Value = versie2;
			rsInstNew.Fields["STATUS_MONTAGE_INSTRUCTIE"].Value = 3;
			rsInstNew.Fields["OMSCHRIJVING"].Value = omschrijving;
			rsInstNew.Fields["INSTRUCTIE_TEKST"].Value = instructie;
			rsInstNew.Fields["ONDERDELENLIJST"].Value = onderdelen;			

			rsInstNew.Fields["AFBEELDING_A"].Value = imageA;
			rsInstNew.Fields["AFBEELDING_B"].Value = imageB;
			rsInstNew.Fields["AFBEELDING_C"].Value = imageC;
			rsInstNew.Fields["AFBEELDING_D"].Value = imageD;
			rsInstNew.Fields["AFBEELDING_E"].Value = imageE;
			rsInstNew.Fields["AFBEELDING_F"].Value = imageF;
			rsInstNew.Fields["AFBEELDING_G"].Value = imageG;
			rsInstNew.Fields["AFBEELDING_H"].Value = imageH;			
			
			rsInstNew.Update();

		}

		MessageBox.Show("Geselecteerde instructies zijn gereviseerd");


	}
}