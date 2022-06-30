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
	
	Creeer maakdeel, het  programma om een maakdeel te maken uit voorraad artikelen.
	Uit te voeren vanuit een atikel welke opgebouwd is als maakdeel met een stuklijst
	Geschreven door: Machiel R. van Emden mei-2022

	*/
	
	private static DialogResult ShowInputDialog(ref decimal input1)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Leonardo da Vinci";

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "Accept";
		okButton.Text = "&OK";
		okButton.Size = new System.Drawing.Size(75, 25);
		okButton.Location = new System.Drawing.Point(5, 10);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "ABORT";
		cancelButton.Text = "&Cancel";
		cancelButton.Size = new System.Drawing.Size(75, 25);
		cancelButton.Location = new System.Drawing.Point(100, 10);
		inputBox.Controls.Add(cancelButton);


		//groep prijs
		GroupBox groepprijs = new GroupBox();
		groepprijs.Size = new System.Drawing.Size(180, 60);
		groepprijs.Location = new System.Drawing.Point(10, 75);
		groepprijs.Text = "Aantal te maken";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1000;
		textBox1.DecimalPlaces = 0;
		groepprijs.Controls.Add(textBox1);

		inputBox.Controls.Add(groepprijs);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;


		DialogResult result = inputBox.ShowDialog();

		input1 = textBox1.Value;

		return result;
	}

	public void Execute()
	{
		decimal input1 = 0;
		ShowInputDialog(ref input1);

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		if (input1 == 0)
		{
			MessageBox.Show("Aantal mag geen 0 zijn");
			return;
		}


		List<string> UitLijst = new List<string>();
		List<string> FoutLijst = new List<string>();

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", "PK_R_ITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();

			string ItemNmr = rsItem.Fields["PK_R_ITEM"].Value.ToString();

			ScriptRecordset rsStuklijst = this.GetRecordset("R_ASSEMBLY", "", "FK_ITEM = " + ItemNmr, "");
			rsStuklijst.MoveFirst();

			if (rsStuklijst.RecordCount == 0)
			{
				MessageBox.Show("artikel is geen maakdeel");
				return;
			}
			string StuklijstId = rsStuklijst.Fields["PK_R_ASSEMBLY"].Value.ToString();

			// lijst met benodigde artikelen maken
			ScriptRecordset rsSlRegel = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "FK_ASSEMBLY = " + StuklijstId, "");
			rsSlRegel.MoveFirst();


			// loop om aanwezige voorraad te checken
			rsSlRegel.MoveFirst();
			while (rsSlRegel.EOF == false)
			{
				string Item1 = rsSlRegel.Fields["FK_ITEM"].Value.ToString();
				decimal AantalNodig = input1 * Convert.ToInt32(rsSlRegel.Fields["QUANTITY"].Value.ToString());
				ScriptRecordset rsItemCheck1 = this.GetRecordset("R_ITEM", "", "PK_R_ITEM = " + Item1, "");
				rsItemCheck1.MoveFirst();

				string Item1Code = rsItemCheck1.Fields["CODE"].Value.ToString();
				string Item1Naam = rsItemCheck1.Fields["DESCRIPTION"].Value.ToString();
				string Item1Vtype = rsItemCheck1.Fields["STOCKLINKTYPE"].Value.ToString();
				string Item1VIn = rsItemCheck1.Fields["TOTALSTOCKIN"].Value.ToString();
				string Item1VOut = rsItemCheck1.Fields["TOTALSTOCKOUT"].Value.ToString();
				string Item1VVast = rsItemCheck1.Fields["TOTALSTOCKRESERVATION"].Value.ToString();

				int Item1VIn1 = Convert.ToInt32(Item1VIn);
				int Item1VOut1 = Convert.ToInt32(Item1VOut);
				int Item1VVast1 = Convert.ToInt32(Item1VVast);

				int Item1VVrij = Item1VIn1 - Item1VOut1 - Item1VVast1;

				if (AantalNodig > Item1VVrij)
				{
					FoutLijst.Add(AantalNodig + "x - " + Item1Code + " - " + Item1Naam + " - te weinig voorraad");
				}

				if (Item1Vtype == "1")
				{
					FoutLijst.Add(AantalNodig + "x - " + Item1Code + " - " + Item1Naam + " - artikel niet voorraadhoudend");
				}

				rsSlRegel.MoveNext();

			}

			// bericht met de fouten
			if (FoutLijst.Count > 0)
			{
				var message2 = string.Join(Environment.NewLine, FoutLijst);

				MessageBox.Show(message2, "Fouten lijst");
				return;

			}

			// loop om gebruikte artikelen uit te boeken
			rsSlRegel.MoveFirst();
			while (rsSlRegel.EOF == false)
			{

				ScriptRecordset rsArtikelUit = this.GetRecordset("R_STOCKOUT", "", "PK_R_STOCKOUT= -1", "");
				rsArtikelUit.UseDataChanges = true;
				rsArtikelUit.AddNew();

				decimal aantaleruit = input1 * Convert.ToInt32(rsSlRegel.Fields["QUANTITY"].Value.ToString());

				string EruitAantal = aantaleruit.ToString();
				string EruitNaam = rsSlRegel.Fields["DESCRIPTION"].Value.ToString();

				rsArtikelUit.Fields["FK_ITEM"].Value = rsSlRegel.Fields["FK_ITEM"].Value;
				rsArtikelUit.Fields["QUANTITY"].Value = aantaleruit;
				rsArtikelUit.Fields["DESCRIPTION"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();

				rsArtikelUit.Update();

				UitLijst.Add(EruitAantal + "x - " + EruitNaam + " - uit geboekt");

				rsSlRegel.MoveNext();

			}

			// maakdeel aanvullen
			ScriptRecordset rsArtikelIn = this.GetRecordset("R_STOCKIN", "", "PK_R_STOCKIN= -1", "");
			rsArtikelIn.UseDataChanges = true;
			rsArtikelIn.AddNew();

			rsArtikelIn.Fields["FK_ITEM"].Value = (int)record.GetPrimaryKeyValue();
			rsArtikelIn.Fields["QUANTITY"].Value = input1;
			rsArtikelIn.Fields["DESCRIPTION"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();

			rsArtikelIn.Update();
			MessageBox.Show("Voorraad van maakdeel aangevult");

			// resultaat bericht
			var message = string.Join(Environment.NewLine, UitLijst);
			MessageBox.Show(message, "totaal uitgeboekt");

		}


	}



}