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
	Aangepast door Machiel R. van Emden september 2023

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
			ScriptRecordset rsSlRegel = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "FK_ASSEMBLY = " + StuklijstId, "POSITION");
			rsSlRegel.MoveFirst();
			
			List<string> FoutLijst = new List<string>();

			// loop om totaal aanwezige voorraad te checken			
			Totalcheck(ref rsSlRegel, ref input1, ref FoutLijst );
			
			// bericht met de fouten
			if (FoutLijst.Count > 0)
			{
				var message2 = string.Join(Environment.NewLine, FoutLijst);
				MessageBox.Show(message2, "Fouten lijst");
				return;

			}
			
			

			// loop om te checken of er gesplitst moet worden
			int checkert = 0;
			Splitcheck(ref rsSlRegel, ref input1, ref checkert );
			
			
			

			if (checkert == 0) // uit en in boeken zonder te splitten
			{
				// loop om gebruikte artikelen uit te boeken zonder splitsen
				TotalRemove(ref rsSlRegel, ref rsItem, ref input1);

				// maakdeel aanvullen
				AddNew(ref rsItem, ref ItemNmr, ref input1);
			}
			
			

			else if (checkert > 0) //uit en in boeken met splitten
			{
				// loop om gebruikte artikelen uit te boeken met splitsen
				PartRemove(ref rsSlRegel, ref rsItem, ref input1);

				// maakdeel aanvullen
				AddNew(ref rsItem, ref ItemNmr, ref input1);
			}			
			

			else // error
			{
				MessageBox.Show("Geen idee");
				return;
			}		
			

			MessageBox.Show("Done");


		}
	}

	public void Totalcheck(ref ScriptRecordset rsSlRegel, ref decimal input1, ref List<string> FoutLijst )
	{
		
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

		
	}

	public void Splitcheck(ref ScriptRecordset rsSlRegel, ref decimal input1, ref int checkert)
	{
		List<string> SplitLijst = new List<string>();
		
		rsSlRegel.MoveFirst();
		while (rsSlRegel.EOF == false)
		{
			string Item1 = rsSlRegel.Fields["FK_ITEM"].Value.ToString();
			decimal AantalNodig = input1 * Convert.ToInt32(rsSlRegel.Fields["QUANTITY"].Value.ToString());

			ScriptRecordset rsItemIn1 = this.GetRecordset("R_STOCKIN", "", "FK_ITEM = " + Item1, "");
			rsItemIn1.MoveFirst();

			while (rsItemIn1.EOF == false)
			{
				string stockinID = rsItemIn1.Fields["PK_R_STOCKIN"].Value.ToString();
				int stockinNumber = Convert.ToInt32(rsItemIn1.Fields["QUANTITY"].Value.ToString());
				int stockuitNumber = 0;
				int stockReservNumber = 0;

				ScriptRecordset rsItemUit1 = this.GetRecordset("R_STOCKOUT", "", "FK_STOCKIN = " + stockinID, "");
				rsItemUit1.MoveFirst();				

				while (rsItemUit1.EOF == false)
				{
					int stockuit = Convert.ToInt32(rsItemUit1.Fields["QUANTITY"].Value.ToString());
					stockuitNumber += stockuit;					
					rsItemUit1.MoveNext();
				}

				ScriptRecordset rsItemReserv1 = this.GetRecordset("R_ITEMRESERVATION", "", "FK_STOCKIN = " + stockinID, "");
				rsItemReserv1.MoveFirst();

				while (rsItemReserv1.EOF == false)
				{
					int stockreserv = Convert.ToInt32(rsItemReserv1.Fields["QUANTITY"].Value.ToString());
					stockReservNumber += stockreserv;
					rsItemReserv1.MoveNext();
				}				

				int stockfree = stockinNumber - stockuitNumber- stockReservNumber;
				

				if (stockfree >= AantalNodig)
				{
					break;				
				
				}

				else 
				{
					rsItemIn1.MoveNext();

					if (rsItemIn1.EOF == true)
					{
						checkert += 1;
						// SplitLijst.Add(Item1 + " moet gesplitst worden");
					}					
					
				}				
								
			}		


			rsSlRegel.MoveNext();


		}
		

			// resultaat bericht
			var message = string.Join(Environment.NewLine, SplitLijst);
			MessageBox.Show(message, "Nodig te splitten");


		
	}	

	public void TotalRemove(ref ScriptRecordset rsSlRegel, ref ScriptRecordset rsItem, ref decimal input1)
	{
		List<string> UitLijst = new List<string>();
		
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
			rsArtikelUit.Fields["MEMO"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();

			rsArtikelUit.Update();

			UitLijst.Add(EruitAantal + "x - " + EruitNaam + " - uit geboekt");

			rsSlRegel.MoveNext();

		}

		// resultaat bericht
		var message = string.Join(Environment.NewLine, UitLijst);
		MessageBox.Show(message, "Totaal uitgeboekt");
	}

	public void PartRemove(ref ScriptRecordset rsSlRegel, ref ScriptRecordset rsItem, ref decimal input1)
	{
		List<string> UitLijst = new List<string>();
		
		rsSlRegel.MoveFirst();
		while (rsSlRegel.EOF == false)
		{
			string Item1 = rsSlRegel.Fields["FK_ITEM"].Value.ToString();
			decimal AantalNodig = input1 * Convert.ToInt32(rsSlRegel.Fields["QUANTITY"].Value.ToString());

			ScriptRecordset rsItemIn1 = this.GetRecordset("R_STOCKIN", "", "FK_ITEM = " + Item1, "");
			rsItemIn1.MoveFirst();

			while (rsItemIn1.EOF == false)
			{
				string stockinID = rsItemIn1.Fields["PK_R_STOCKIN"].Value.ToString();
				int stockinNumber = Convert.ToInt32(rsItemIn1.Fields["QUANTITY"].Value.ToString());
				int stockuitNumber = 0;
				int stockReservNumber = 0;

				ScriptRecordset rsItemUit1 = this.GetRecordset("R_STOCKOUT", "", "FK_STOCKIN = " + stockinID, "");
				rsItemUit1.MoveFirst();

				while (rsItemUit1.EOF == false)
				{
					int stockuit = Convert.ToInt32(rsItemUit1.Fields["QUANTITY"].Value.ToString());
					stockuitNumber += stockuit;
					rsItemUit1.MoveNext();
				}

				ScriptRecordset rsItemReserv1 = this.GetRecordset("R_ITEMRESERVATION", "", "FK_STOCKIN = " + stockinID, "");
				rsItemReserv1.MoveFirst();

				while (rsItemReserv1.EOF == false)
				{
					int stockreserv = Convert.ToInt32(rsItemReserv1.Fields["QUANTITY"].Value.ToString());
					stockReservNumber += stockreserv;
					rsItemReserv1.MoveNext();
				}		

				int stockfree = stockinNumber - stockuitNumber - stockReservNumber;

				if (stockfree >= AantalNodig)
				{
					ScriptRecordset rsArtikelUit = this.GetRecordset("R_STOCKOUT", "", "PK_R_STOCKOUT= -1", "");
					rsArtikelUit.UseDataChanges = true;
					rsArtikelUit.AddNew();

					decimal aantaleruit = AantalNodig;
					string EruitAantal = aantaleruit.ToString();
					string EruitNaam = rsSlRegel.Fields["DESCRIPTION"].Value.ToString();

					rsArtikelUit.Fields["FK_ITEM"].Value = rsSlRegel.Fields["FK_ITEM"].Value;
					rsArtikelUit.Fields["QUANTITY"].Value = aantaleruit;
					rsArtikelUit.Fields["DESCRIPTION"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();
					rsArtikelUit.Fields["MEMO"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();
					rsArtikelUit.Fields["FK_STOCKIN"].Value = rsItemIn1.Fields["PK_R_STOCKIN"].Value;

					rsArtikelUit.Update();

					UitLijst.Add(EruitAantal + "x - " + EruitNaam + " - uit geboekt");
					
					break;
				}
				
				

				else if (stockfree < AantalNodig && stockfree !=0)
				{
					ScriptRecordset rsArtikelUit = this.GetRecordset("R_STOCKOUT", "", "PK_R_STOCKOUT= -1", "");
					rsArtikelUit.UseDataChanges = true;
					rsArtikelUit.AddNew();

					decimal aantaleruit = stockfree;
					string EruitAantal = aantaleruit.ToString();
					string EruitNaam = rsSlRegel.Fields["DESCRIPTION"].Value.ToString();

					rsArtikelUit.Fields["FK_ITEM"].Value = rsSlRegel.Fields["FK_ITEM"].Value;
					rsArtikelUit.Fields["QUANTITY"].Value = aantaleruit;
					rsArtikelUit.Fields["DESCRIPTION"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();
					rsArtikelUit.Fields["MEMO"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();
					rsArtikelUit.Fields["FK_STOCKIN"].Value = rsItemIn1.Fields["PK_R_STOCKIN"].Value;
					
					rsArtikelUit.Update();

					AantalNodig -= stockfree;

					UitLijst.Add(EruitAantal + "x - " + EruitNaam + " - uit geboekt");				
					
					
					rsItemIn1.MoveNext();

				}
				
				else rsItemIn1.MoveNext();



			}




			rsSlRegel.MoveNext();


		}

		// resultaat bericht
		var message = string.Join(Environment.NewLine, UitLijst);
		MessageBox.Show(message, "Totaal uitgeboekt");
		
	}	

	public void AddNew(ref ScriptRecordset rsItem, ref string ItemNmr, ref decimal input1)
	{
		ScriptRecordset rsArtikelIn = this.GetRecordset("R_STOCKIN", "", "PK_R_STOCKIN= -1", "");
		rsArtikelIn.UseDataChanges = true;
		rsArtikelIn.AddNew();

		rsArtikelIn.Fields["FK_ITEM"].Value = Convert.ToInt32(ItemNmr);
		rsArtikelIn.Fields["QUANTITY"].Value = input1;
		rsArtikelIn.Fields["DESCRIPTION"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();
		rsArtikelIn.Fields["MEMO"].Value = "MvE maakdeel script: " + rsItem.Fields["CODE"].Value.ToString() + " - " + rsItem.Fields["DESCRIPTION"].Value.ToString();

		rsArtikelIn.Update();
		MessageBox.Show("Voorraad van maakdeel aangevult");
	}
	

	// M.R.v.E - 2023

}