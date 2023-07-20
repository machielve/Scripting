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
	
	Diesel toeslag verdelen, het  programma om de ingegegven dieseltoeslag te verdelen over de geselecteerde factuur regels
	Uit te voeren vanuit een inkooporderfactuur op niet gejouranliseerde regels
	Geschreven door: Machiel R. van Emden juli-2023

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
		groepprijs.Text = "Totale diesel opslag";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1500000;
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

		decimal input1 = 1;

		ShowInputDialog(ref input1);

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		decimal totaal = 0;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_PURCHASEINVOICEDETAILMISC", "", "PK_R_PURCHASEINVOICEDETAILMISC= " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();

			decimal aantal = Convert.ToDecimal(rsItem.Fields["QUANTITY"].Value.ToString());

			totaal += aantal;

		}

		decimal kgopslag = input1 / totaal;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem1 = this.GetRecordset("R_PURCHASEORDERDETAILMISC", "", "PK_R_PURCHASEORDERDETAILMISC = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem1.MoveFirst();
			rsItem1.UseDataChanges = true;

			decimal huidig = Convert.ToDecimal(rsItem1.Fields["NETPURCHASEPRICE"].Value.ToString());
			decimal aantal1 = Convert.ToDecimal(rsItem1.Fields["QUANTITY"].Value.ToString());

			decimal Opslag = aantal1 * kgopslag;
			decimal Nieuw = huidig + Opslag;

			rsItem1.Fields["NETPURCHASEPRICE"].Value = Nieuw;

			rsItem1.Update();



		}

	}

	// M.R.v.E - 2023

}
