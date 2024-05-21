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
	//	public myForm(){
	private static DialogResult ShowInputDialog(ref decimal input1)
	{
		System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
		customCulture.NumberFormat.NumberDecimalSeparator = ",";

		System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Hattin";

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
		groepprijs.Text = "Prijs / kg";

		System.Windows.Forms.NumericUpDown numericUpDown1 = new NumericUpDown();
		numericUpDown1.Size = new System.Drawing.Size(100, 25);
		numericUpDown1.Location = new System.Drawing.Point(5, 25);
		numericUpDown1.Value = input1;
		numericUpDown1.Minimum = 0;
		numericUpDown1.Maximum = 1500000;
		numericUpDown1.DecimalPlaces = 2;
		numericUpDown1.Controls[0].Visible = false;
		groepprijs.Controls.Add(numericUpDown1);

		inputBox.Controls.Add(groepprijs);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		input1 = numericUpDown1.Value;

		return result;
	}
	
	
	
	
	
	
	public void Execute()
	{
		decimal input1 = 0;
		DialogResult result = ShowInputDialog(ref input1);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Afwijkende kiloprijs afgebroken");
			return;
		}

		decimal kgprice = input1;
		
		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			int aNummer = (int)record.GetPrimaryKeyValue();
			
			ScriptRecordset rsItems = this.GetRecordset("R_ITEM", "", "PK_R_ITEM = " + aNummer, "");
			rsItems.MoveFirst();

			string naam = rsItems.Fields["DESCRIPTION"].Value.ToString();
			decimal gewicht = Convert.ToDecimal(rsItems.Fields["WEIGHT"].Value.ToString());
			string pricename = @"'8b85c821-4b2a-4c6f-96e1-d5202355dd6a'";
			string datum = DateTime.Now.ToShortDateString();
		
			
			ScriptRecordset rsItemSup = this.GetRecordset("R_ITEMSUPPLIER", "", "FK_ITEM = " + aNummer, "");

			if (rsItems.RecordCount == 0)
			{
				rsItemSup.AddNew();

				rsItemSup.Fields["FK_RELATION"].Value = 548;
				rsItemSup.Fields["FK_ITEM"].Value = aNummer;
				rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam;
				rsItemSup.Fields["ITEMTYPE"].Value = 8;
				rsItemSup.Update();
			}
			rsItemSup.MoveFirst();

			while (rsItemSup.EOF == false)
			{
				if (rsItemSup.Fields["FK_RELATION"].Value.ToString() == "548")
				{
					rsItemSup.Fields["MAINSUPPLIER"].Value = 1;
					rsItemSup.Fields["ITEMTYPE"].Value = 8;
					int Itemsup = Convert.ToInt32(rsItemSup.Fields["PK_R_ITEMSUPPLIER"].Value.ToString());

					ScriptRecordset rsItemPrice = this.GetRecordset("R_ITEMPURCHASEPRICE", "", "FK_ITEMSUPPLIER = " + Itemsup + " AND FK_PRICENAME = " + pricename, "");
					rsItemPrice.MoveFirst();

					decimal newprice = gewicht * kgprice;

					rsItemPrice.Fields["VALUE"].Value = newprice;

					string message = "Prijs voor Kampstaal berekend op: " + datum + " met: " + kgprice.ToString() + " €/kg.";

					rsItems.Fields["MEMO"].Value = message;

					rsItems.Update();
					rsItemPrice.Update();
					rsItemSup.Update();
					
					
					break;
				}
				
				else rsItemSup.MoveNext();
				
			}
		}
		
		
		
		
		
		



	}
}