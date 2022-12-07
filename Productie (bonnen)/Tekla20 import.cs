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

		string clipboardData = Clipboard.GetText();

		List<string> errorList = new List<string>();

		foreach (var myString in clipboardData.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
		{
			string[] myStrValues = myString.Split('\t');
			string ItemCode = myStrValues[0].ToString();

			ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", ItemCode), "");
			rsItem.MoveFirst();

			decimal type = Convert.ToDecimal(rsItem.Fields["FK_ITEMUNIT"].Value.ToString());
			string omschrijf = rsItem.Fields["DESCRIPTION"].Value.ToString();

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
				decimal breed;
				decimal lang;
				
				// Artikleeenheden Plaat en Rooster, lengte en breedte
				if (type == 10 || type == 15 || type == 30)
				{
					if (myStrValues[4].ToString() == "")
					{
						errorList.Add(ItemCode + " - " + omschrijf + " - heeft een breedte nodig");
						continue;
					}
					else if (myStrValues[3].ToString() == "")
					{
						errorList.Add(ItemCode + " - " + omschrijf + " - heeft een lengte nodig");
						continue;
					}
					else
					{

						breed = (Convert.ToDecimal(myStrValues[4])) / 1000;
						lang = (Convert.ToDecimal(myStrValues[3])) / 1000;
					}
				}
				
				
				// Artikleeenheden met een lengte maat
				else if (type == 11 || type == 17 || type == 20 || type == 23 || type == 24 || type == 31 || type == 32)
				{
					if (myStrValues[3].ToString() == "")
					{
						errorList.Add(ItemCode + " - " + omschrijf + " - heeft een lengte nodig");
						break;
					}
					else
					{
						breed = 0;
						lang = (Convert.ToDecimal(myStrValues[3])) / 1000;
					}

				}
				
				// Artikleeenheid Trapboom
				else if (type == 22 || type == 34)
				{
					breed = 0;
					lang = (Convert.ToDecimal(myStrValues[3]));
				}
				
				
				// Artikleenheden welke nog niet gebruikt zijn
				else
				{
					breed = 0;
					lang = 0;
				}


				ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
				rsJoborderItem.UseDataChanges = true;
				rsJoborderItem.AddNew();
				
				

				rsJoborderItem.Fields["FK_JOBORDER"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
				rsJoborderItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
				rsJoborderItem.Fields["QUANTITY"].Value = Convert.ToDouble(myStrValues[1]);
				rsJoborderItem.Fields["CAMPARAMETER"].Value = Convert.ToString(myStrValues[2]);
				rsJoborderItem.Fields["LENGTH"].Value = Convert.ToDouble(lang);
				rsJoborderItem.Fields["WIDTH"].Value = Convert.ToDouble(breed);
				rsJoborderItem.Fields["CAMGEOMETRY"].Value = Convert.ToString(myStrValues[5]);				
				rsJoborderItem.Fields["AFWERKING"].Value = Convert.ToString(myStrValues[6]);					

				rsJoborderItem.Update();

			}
		}

		var message = string.Join(Environment.NewLine, errorList);
		MessageBox.Show(message,"Aandachtslijst");



	}
}