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
	
	Meer prijs verdelen, het  programma om de ingegegeven prijs te vergelijken en de meerprijs te verdelen over de geselecteerde factuur regels
	Uit te voeren vanuit een inkooporderfactuur op niet gejouranliseerde regels
	Geschreven door: Machiel R. van Emden april-2024

	*/

	private static DialogResult ShowInputDialog(ref decimal input1)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Minas Tirith";

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
		groepprijs.Text = "Totale factur prijs";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);

		textBox1.DecimalPlaces = 2;
		textBox1.Minimum = -20000;
		textBox1.Maximum = 1500000;
		textBox1.Value = input1;
		textBox1.DecimalPlaces = 2;
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

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		decimal totaal = 0; // Totaal prijs
		decimal input1 = 0; // Totaal gewicht

		foreach (IRecord record in records)
		{
			ScriptRecordset rsInvoice = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "", "PK_R_PURCHASEINVOICEDETAILITEM= " + (int)record.GetPrimaryKeyValue(), "");
			rsInvoice.MoveFirst();

			decimal aantal = Convert.ToDecimal(rsInvoice.Fields["NETPURCHASEPRICE"].Value.ToString());

			input1 += aantal;
		}

		foreach (IRecord record in records)
		{
			ScriptRecordset rsInvoice = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "", "PK_R_PURCHASEINVOICEDETAILITEM= " + (int)record.GetPrimaryKeyValue(), "");
			rsInvoice.MoveFirst();

			int ontvangst = Convert.ToInt32(rsInvoice.Fields["FK_GOODSRECEIPTDETAILITEM"].Value.ToString());

			ScriptRecordset rsReceive = this.GetRecordset("R_GOODSRECEIPTDETAILITEM", "", "PK_R_GOODSRECEIPTDETAILITEM= " + ontvangst, "");
			rsReceive.MoveFirst();

			decimal aantal = Convert.ToDecimal(rsReceive.Fields["NETPURCHASEPRICE"].Value.ToString());

			totaal += aantal;
		}
		
		decimal percentageOpslag = input1 / totaal;
		
		

		DialogResult result = ShowInputDialog(ref input1);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Factuur afwijking afgebroken");
			return;
		}

		

		


		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem1 = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "", "PK_R_PURCHASEINVOICEDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem1.MoveFirst();
			rsItem1.UseDataChanges = true;
			
			int ontvangst = Convert.ToInt32(rsItem1.Fields["FK_GOODSRECEIPTDETAILITEM"].Value.ToString());

			ScriptRecordset rsReceive = this.GetRecordset("R_GOODSRECEIPTDETAILITEM", "", "PK_R_GOODSRECEIPTDETAILITEM= " + ontvangst, "");
			rsReceive.MoveFirst();

			decimal huidig = Convert.ToDecimal(rsReceive.Fields["NETPURCHASEPRICE"].Value.ToString());

			decimal Nieuw = huidig * percentageOpslag;

			rsItem1.Fields["NETPURCHASEPRICE"].Value = Nieuw;

			rsItem1.Update();


		}

		MessageBox.Show("Klaar");
		

	}

	// M.R.v.E - 2024

}
