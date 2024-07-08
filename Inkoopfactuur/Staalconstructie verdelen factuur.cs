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
	
	Staalconstuctie verdelen factuur, het  programma om een totaal inkoop bedrag te verdelen per kg over de geselecteerde factuur regels
	Uit te voeren vanuit een inkooporder op niet ontvangen regels
	Geschreven door: Machiel R. van Emden juli-2023

	*/

	private static DialogResult ShowInputDialog(ref decimal input1, ref decimal input4)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Helmsdeep";

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


		//groep staal
		GroupBox groepstaal = new GroupBox();
		groepstaal.Size = new System.Drawing.Size(180, 60);
		groepstaal.Location = new System.Drawing.Point(10, 75);
		groepstaal.Text = "Project totaal prijs";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1500000;
		textBox1.DecimalPlaces = 2;
		groepstaal.Controls.Add(textBox1);

		inputBox.Controls.Add(groepstaal);

		//groep transport
		GroupBox groeptransport = new GroupBox();
		groeptransport.Size = new System.Drawing.Size(180, 60);
		groeptransport.Location = new System.Drawing.Point(10, 300);
		groeptransport.Text = "Transport totaal prijs";

		System.Windows.Forms.NumericUpDown textBox4 = new NumericUpDown();
		textBox4.Size = new System.Drawing.Size(100, 25);
		textBox4.Location = new System.Drawing.Point(5, 25);
		textBox4.Value = input4;
		textBox4.Minimum = 0;
		textBox4.Maximum = 1500000;
		textBox4.DecimalPlaces = 2;
		groeptransport.Controls.Add(textBox4);

		inputBox.Controls.Add(groeptransport);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		input1 = textBox1.Value;

		input4 = textBox4.Value;

		return result;
	}

	public void Execute()
	{

		decimal input1 = 1; // Totaal prijs
		decimal input4 = 0; // Totaal Transport
		decimal totaal = 0; // Totaal gewicht

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItemI = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "PK_R_PURCHASEINVOICEDETAILITEM", "PK_R_PURCHASEINVOICEDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItemI.MoveFirst();
			string IRegelID = rsItemI.Fields["PK_R_PURCHASEINVOICEDETAILITEM"].Value.ToString();
			
			decimal prijs = Convert.ToDecimal(rsItemI.Fields["NETPURCHASEPRICE"].Value);

			ScriptRecordset rsItemR = this.GetRecordset("R_ITEMRESERVATION", "DESTINATIONID", "FK_PURCHASEINVOICEDETAILITEM = " + IRegelID, "");
			rsItemR.MoveFirst();
			string RRegelID = rsItemR.Fields["DESTINATIONID"].Value.ToString();

			ScriptRecordset rsItemB = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= " + RRegelID, "");
			rsItemB.MoveFirst();
			decimal gewicht = Convert.ToDecimal(rsItemB.Fields["WEIGHT"].Value);

			totaal += gewicht;
			input1 += prijs;
		}
		

		DialogResult result1 = ShowInputDialog(ref input1, ref input4);

		if (result1 != DialogResult.OK)
		{
			MessageBox.Show("Staalconstructie verdelen afgebroken");
			return;
		}

		decimal input11 = Math.Round(input1, 2);

		decimal result = input1 / totaal;

		MessageBox.Show("Totaal € " + Math.Round(result, 2) + " / kg");


		foreach (IRecord record in records)
		{
			ScriptRecordset rsItemI = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "", "PK_R_PURCHASEINVOICEDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItemI.MoveFirst();
			rsItemI.UseDataChanges = true;
			string IRegelID = rsItemI.Fields["PK_R_PURCHASEINVOICEDETAILITEM"].Value.ToString();


			ScriptRecordset rsItemR = this.GetRecordset("R_ITEMRESERVATION", "DESTINATIONID", "FK_PURCHASEINVOICEDETAILITEM = " + IRegelID, "");
			rsItemR.MoveFirst();
			string RRegelID = rsItemR.Fields["DESTINATIONID"].Value.ToString();


			ScriptRecordset rsItemB = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= " + RRegelID, "");
			rsItemB.MoveFirst();
			decimal gewicht = Convert.ToDecimal(rsItemB.Fields["WEIGHT"].Value);
			decimal prijs = gewicht * result;

			rsItemI.Fields["NETPURCHASEPRICE"].Value = prijs;

			rsItemI.Update();

		}

	}

	// M.R.v.E - 2023

}
