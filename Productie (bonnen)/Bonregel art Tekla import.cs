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
		//MvE - 28-7-2020: Plakken artikelregels obv excel lijst in Productiebon.
		//Laatste update 28-7-2020
		//Opbouw: Posnummer/ Artikelcode/ Omschrijving/ Aantal/ Memo

		string clipboardData = Clipboard.GetText();

		foreach (var myString in clipboardData.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
		{
			string[] myStrValues = myString.Split('\t');
			string ItemCode = myStrValues[0].ToString();

			ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", ItemCode), "");
			rsItem.MoveFirst();

			if (rsItem != null && rsItem.RecordCount == 0)
			{
				MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + ItemCode);
			}
			else if (myStrValues[0].ToString() == "")
			{
				MessageBox.Show("Artikel: " + ItemCode + " heeft geen aantal ingevuld.");
			}


			else
			{
				ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
				rsJoborderItem.UseDataChanges = true;
				rsJoborderItem.AddNew();

				rsJoborderItem.Fields["FK_JOBORDER"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
				rsJoborderItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
				rsJoborderItem.Fields["QUANTITY"].Value = Convert.ToDouble(myStrValues[1]);
				rsJoborderItem.Fields["CAMPARAMETER"].Value = Convert.ToString(myStrValues[2]);
				rsJoborderItem.Fields["LENGTH"].Value = Convert.ToDouble(myStrValues[3]);
				//rsJoborderItem.Fields["WEIGHT"].Value = Convert.ToDouble(myStrValues[4]);

				rsJoborderItem.Update();

			}
		}
		

	}
}