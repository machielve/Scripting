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
	
	Vloerprijs verdelen, het  programma om een totaal inkoop bedrag te verdelen per m² over de geselecteerde regels
	Uit te voeren vanuit een inkooporder op niet ontvangen regels
	Geschreven door: Machiel R. van Emden mei-2022

	*/
	
	private static DialogResult ShowInputDialog(ref decimal input1, ref decimal input2, ref decimal input3, ref decimal input4)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Stirling bridge";

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


		//groep hout
		GroupBox groephout = new GroupBox();
		groephout.Size = new System.Drawing.Size(180, 60);
		groephout.Location = new System.Drawing.Point(10, 75);
		groephout.Text = "Hout totaal prijs";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1500000;
		textBox1.DecimalPlaces = 2;
		groephout.Controls.Add(textBox1);

		inputBox.Controls.Add(groephout);

		//groep zagen
		GroupBox groepzaag = new GroupBox();
		groepzaag.Size = new System.Drawing.Size(180, 60);
		groepzaag.Location = new System.Drawing.Point(10, 150);
		groepzaag.Text = "Zagen totaal prijs";

		System.Windows.Forms.NumericUpDown textBox2 = new NumericUpDown();
		textBox2.Size = new System.Drawing.Size(100, 25);
		textBox2.Location = new System.Drawing.Point(5, 25);
		textBox2.Value = input2;
		textBox2.Minimum = 0;
		textBox2.Maximum = 1500000;
		textBox2.DecimalPlaces = 2;
		groepzaag.Controls.Add(textBox2);

		inputBox.Controls.Add(groepzaag);

		//groep groeven
		GroupBox groepgroef = new GroupBox();
		groepgroef.Size = new System.Drawing.Size(180, 60);
		groepgroef.Location = new System.Drawing.Point(10, 225);
		groepgroef.Text = "Groeven totaal prijs";

		System.Windows.Forms.NumericUpDown textBox3 = new NumericUpDown();
		textBox3.Size = new System.Drawing.Size(100, 25);
		textBox3.Location = new System.Drawing.Point(5, 25);
		textBox3.Value = input3;
		textBox3.Minimum = 0;
		textBox3.Maximum = 1500000;
		textBox3.DecimalPlaces = 2;
		groepgroef.Controls.Add(textBox3);

		inputBox.Controls.Add(groepgroef);

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
		input2 = textBox2.Value;
		input3 = textBox3.Value;
		input4 = textBox4.Value;



		return result;
	}

	public void Execute()
	{

		decimal input1 = 1;
		decimal input2 = 1;
		decimal input3 = 1;
		decimal input4 = 1;

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();
		
		string InkoopNummer = this.FormDataAwareFunctions.FormParent.CurrentRecord.GetPrimaryKeyValue().ToString();

	//	MessageBox.Show(InkoopNummer);

		if (records.Length == 0)
			return;

		decimal totaal = 0;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_PURCHASEORDERDETAILITEM", "LENGTH, WIDTH, QUANTITY", "PK_R_PURCHASEORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			decimal lengte = Convert.ToDecimal(rsItem.Fields["LENGTH"].Value);
			decimal breedte = Convert.ToDecimal(rsItem.Fields["WIDTH"].Value);
			decimal aantal = Convert.ToDecimal(rsItem.Fields["QUANTITY"].Value);

			decimal opp = lengte * breedte * aantal;

			decimal output4 = opp;

			totaal += output4;


		}


		decimal input6 = totaal;

		ShowInputDialog(ref input1, ref input2, ref input3, ref input4);

		decimal output1 = input1 + input2 + input3;

		decimal input11 = Math.Round(input1, 2);
		decimal input21 = Math.Round(input2, 2);
		decimal input31 = Math.Round(input3, 2);
		decimal input41 = Math.Round(input4, 2);

		decimal output2 = output1 / input6;

		decimal output3 = Math.Round(output2, 2);


		MessageBox.Show("€ " + input11 + " totale houtprijs" +
						"\n" + "€ " + input21 + " totale zaagprijs" +
						"\n" + "€ " + input31 + " totale groefprijs" +
						"\n" + "-----------------------" +
						"\n" + "€ " + output1 + " totale bewerkte prijs" +
						"\n" + input6 + " m² benodigd" +
						"\n" + "-----------------------" +
						"\n" + "€ " + output3 + " /m² bruto inkoopprijs" +
						"\n" +
						"\n" + "-----------------------" +
						"\n" + "€ " + input41 + " totale transportprijs" +
						"\n" +

							"\nBedrag word nu aangepast in inkooporder :)", "Balista");



		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_PURCHASEORDERDETAILITEM", "", "PK_R_PURCHASEORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			rsItem.Fields["GROSSPURCHASEPRICE"].Value = output3;

			rsItem.Update();

		}


		decimal test = 0;
		decimal totaal2 = 0;
		string test1 = "0";

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_PURCHASEORDERDETAILITEM", "NETPURCHASEPRICE,PK_R_PURCHASEORDERDETAILITEM", "PK_R_PURCHASEORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			decimal prijs = Convert.ToDecimal(rsItem.Fields["NETPURCHASEPRICE"].Value);

			totaal2 += prijs;

			if (prijs > test)
			{
				test = prijs;
				test1 = rsItem.Fields["PK_R_PURCHASEORDERDETAILITEM"].Value.ToString();
			}

		}

		decimal verschil = output1 - totaal2;

		ScriptRecordset rsItem9 = this.GetRecordset("R_PURCHASEORDERDETAILITEM", "", "PK_R_PURCHASEORDERDETAILITEM = " + test1, "");
		rsItem9.MoveFirst();
		rsItem9.UseDataChanges = true;

		rsItem9.Fields["NETPURCHASEPRICE"].Value = Convert.ToDecimal(rsItem9.Fields["NETPURCHASEPRICE"].Value) + verschil;

		rsItem9.Update();


		decimal test50 = 0;
		decimal totaal51 = 0;


		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_PURCHASEORDERDETAILITEM", "NETPURCHASEPRICE", "PK_R_PURCHASEORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			decimal prijs = Convert.ToDecimal(rsItem.Fields["NETPURCHASEPRICE"].Value);

			totaal51 += prijs;

			if (prijs > test50)
			{
				test50 = prijs;

			}
			
			

		}


		ScriptRecordset rsTransport = this.GetRecordset("R_PURCHASEORDERDETAILMISC", "", "FK_PURCHASEORDER = " + Convert.ToInt32(InkoopNummer), "");
		rsTransport.MoveFirst();
		rsTransport.UseDataChanges = true;

		rsTransport.Fields["NETPURCHASEPRICE"].Value = input41;

		rsTransport.Update();



	}


}
