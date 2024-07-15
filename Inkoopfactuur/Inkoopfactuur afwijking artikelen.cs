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

	private static DialogResult ShowInputDialog(ref decimal input1, ref string totaal1)
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


		//groep factuur prijs
		GroupBox groepprijs = new GroupBox();
		groepprijs.Size = new System.Drawing.Size(250, 60);
		groepprijs.Location = new System.Drawing.Point(10, 75);
		groepprijs.Text = "Totale factur prijs artikelen";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(150, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);

		textBox1.DecimalPlaces = 2;
		textBox1.Minimum = -20000;
		textBox1.Maximum = 1500000;
		textBox1.Value = input1;
		textBox1.DecimalPlaces = 2;
		groepprijs.Controls.Add(textBox1);

		inputBox.Controls.Add(groepprijs);

		//groep ontvangst prijs
		GroupBox receiveprijs = new GroupBox();
		receiveprijs.Size = new System.Drawing.Size(250, 90);
		receiveprijs.Location = new System.Drawing.Point(10, 140);
		receiveprijs.Text = "Totale ontvangst prijs artikelen";
		
		System.Windows.Forms.Label label1 = new Label();
		label1.Size = new System.Drawing.Size(200, 25);
		label1.Location = new System.Drawing.Point(5, 25);
		label1.Text = totaal1;
		receiveprijs.Controls.Add(label1);
		
		
		

		inputBox.Controls.Add(receiveprijs);
		
		
		
		
		


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

		decimal totaal = 0; // Totaal ontvangst bedrag
		decimal input1 = 0; // Totaal factuur bedrag
		
		
		// Factuur bedrag uitrekenen
		foreach (IRecord record in records)
		{
			ScriptRecordset rsInvoice = this.GetRecordset("R_PURCHASEINVOICEDETAILITEM", "", "PK_R_PURCHASEINVOICEDETAILITEM= " + (int)record.GetPrimaryKeyValue(), "");
			rsInvoice.MoveFirst();

			decimal aantal = Convert.ToDecimal(rsInvoice.Fields["NETPURCHASEPRICE"].Value.ToString());

			input1 += aantal;
		}

		
		// Ontvangst bedrag uitrekenen
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

		string totaal1 = "totaal ontvangen = € " + Convert.ToString(totaal);


		
		// Afwijking berekenen
		decimal percentageOpslag = input1 / totaal;
		
		

		//pop-up
		DialogResult result = ShowInputDialog(ref input1, ref totaal1);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Factuur afwijking afgebroken");
			return;
		}

		

		


		// Nieuwe regels berekenen
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
