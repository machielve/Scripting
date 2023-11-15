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
	private static DialogResult ShowInputDialog(ref string input)
	{
		System.Drawing.Size size = new System.Drawing.Size(200, 90);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Waterloo";

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
		textBox.Location = new System.Drawing.Point(5, 5);
		textBox.Text = input;
		inputBox.Controls.Add(textBox);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "cancelButton";
		cancelButton.Size = new System.Drawing.Size(75, 23);
		cancelButton.Text = "&Cancel";
		cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();
		input = textBox.Text;
		return result;
	}

	public void Execute()
	{
		string input = "Afwijkende leverancier";
		ShowInputDialog(ref input);

		string input1 = "'" + input + "'";
		

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		ScriptRecordset rsRelation = this.GetRecordset("R_RELATION", "", @"NAME = " + input1, "");
		rsRelation.MoveFirst();

		if (rsRelation.RecordCount == 0)
		{
			MessageBox.Show("Leverancier onbekend");

			return;
		}

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			rsItem.Fields["FK_SUPPLIER"].Value = rsRelation.Fields["PK_R_RELATION"].Value;

			rsItem.Update(null, null);

			/*
			
		
			1 check of de combo artikel + leverancier bestaat
			2 als bestaat dan return
			3 anders aanmaken

			
			*/



			string aCode = rsItem.Fields["FK_ITEM"].Value.ToString();
			string supCode = rsRelation.Fields["PK_R_RELATION"].Value.ToString();
			string inkoopnaam = rsItem.Fields["DESCRIPTION"].Value.ToString();
			string inkoopcode = "";
			int arttype = 1;
			

			int check1 = 0;


			

			ScriptRecordset rsItemSup = this.GetRecordset("R_ITEMSUPPLIER", "", "FK_ITEM = " + aCode, "");
			rsItemSup.MoveFirst();

			if (rsItemSup.RecordCount > 0)

			{
				while (rsItemSup.EOF == false)
				{
					
					if (rsItemSup.Fields["FK_RELATION"].Value.ToString() == supCode)
					{
						check1 = +1;
						rsItemSup.MoveNext();

					}

					else
					{
						rsItemSup.MoveNext();

					}
				}


				if (check1 > 0 )
				{
					return;
				}

				else
				{
					CreateItemSup(ref aCode, ref supCode, ref inkoopnaam, ref inkoopcode, ref arttype);
				}

			}
			
			

			else CreateItemSup(ref aCode, ref supCode, ref inkoopnaam, ref inkoopcode, ref arttype);
			
			rsItem.Update(null, null);


		}

	}

	public void CreateItemSup(	ref string aCode, ref string supCode, ref string inkoopnaam, ref string inkoopcode, ref int arttype )

	{
		//MessageBox.Show("New part supplier");

		// create new item supplier
		ScriptRecordset rsItemS = this.GetRecordset("R_ITEMSUPPLIER", "", "PK_R_ITEMSUPPLIER = -1", "");
		rsItemS.UseDataChanges = true;
		rsItemS.AddNew();

		rsItemS.Fields["FK_ITEM"].Value = Convert.ToInt32(aCode);
		rsItemS.Fields["FK_RELATION"].Value = Convert.ToInt32(supCode);
		rsItemS.Fields["PURCHASEDESCRIPTION"].Value = inkoopnaam;
		rsItemS.Fields["PURCHASEITEMCODE"].Value = inkoopcode;
		rsItemS.Fields["ITEMTYPE"].Value = arttype;
		

		rsItemS.Update(null, null);

		//Create itemsup

	}
}

