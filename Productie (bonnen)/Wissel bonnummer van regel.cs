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

	private static DialogResult ShowInputDialog(ref Decimal input)
	{
		System.Drawing.Size size = new System.Drawing.Size(200, 90);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Bon nummer";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(50, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1500000;
		textBox1.DecimalPlaces = 0;
		inputBox.Controls.Add(textBox1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 50);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "cancelButton";
		cancelButton.Size = new System.Drawing.Size(75, 23);
		cancelButton.Text = "&Cancel";
		cancelButton.Location = new System.Drawing.Point(size.Width - 80, 50);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;




		DialogResult result = inputBox.ShowDialog();
		input = textBox1.Value;
		return result;
	}


	public void Execute()
	{
		/*
		Bon wissel,
		het script om een bonregel artikel door te shuiven naar een andere bon onder dezelfde order.
		alleen voor bonregels die nog geen inkoop er aan gekoppeld hebben.		

		*/

		decimal input = 1;
		ShowInputDialog(ref input);

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
		{
			MessageBox.Show("Geen regels geselecteerd");
			return;
		}

		foreach (IRecord record in records)
		{
			ScriptRecordset rsJoborderItemOld = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= " + (int)record.GetPrimaryKeyValue(), "");
			rsJoborderItemOld.MoveFirst();

			// check of bon bestaat
			int OrderID = Int32.Parse(rsJoborderItemOld.Fields["FK_ORDER"].Value.ToString());
			int Bonchecker = 0;


			ScriptRecordset rsBonList = this.GetRecordset("R_JOBORDER", "", "FK_ORDER= " + OrderID, "JOBORDERNUMBER");
			rsBonList.MoveFirst();


			while (rsBonList.EOF == false)
			{
				if (rsBonList.Fields["JOBORDERNUMBER"].Value.ToString() == input.ToString())
				{
					Bonchecker += Convert.ToInt32(rsBonList.Fields["PK_R_JOBORDER"].Value.ToString());
					rsBonList.MoveNext();
				}
				else rsBonList.MoveNext();
			}

			// toevoegen als bon niet bestaat
			if (Bonchecker == 0)
			{
				MessageBox.Show("Bon bestaat (nog) niet");
				rsBonList.AddNew();

				rsBonList.Fields["FK_ORDER"].Value = OrderID;
				rsBonList.Update();

				Bonchecker += Convert.ToInt32(rsBonList.Fields["PK_R_JOBORDER"].Value.ToString());

				MessageBox.Show("Bon aangemaakt");

			}

			ScriptRecordset rsJoborderItemNew = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
			rsJoborderItemNew.AddNew();

			// nieuwe regels aanmaken
			rsJoborderItemNew.Fields["FK_JOBORDER"].Value = Bonchecker;

			rsJoborderItemNew.Fields["WEIGHT"].Value = rsJoborderItemOld.Fields["WEIGHT"].Value;
			rsJoborderItemNew.Fields["FK_ORDER"].Value = rsJoborderItemOld.Fields["FK_ORDER"].Value;
			rsJoborderItemNew.Fields["FK_ITEMWAREHOUSE"].Value = rsJoborderItemOld.Fields["FK_ITEMWAREHOUSE"].Value;
			rsJoborderItemNew.Fields["DELIVERYMETHOD"].Value = rsJoborderItemOld.Fields["DELIVERYMETHOD"].Value;
			rsJoborderItemNew.Fields["DESCRIPTION"].Value = rsJoborderItemOld.Fields["DESCRIPTION"].Value;
			rsJoborderItemNew.Fields["REGISTRATIONPATH"].Value = rsJoborderItemOld.Fields["REGISTRATIONPATH"].Value;
			rsJoborderItemNew.Fields["SAWINGCODE"].Value = rsJoborderItemOld.Fields["SAWINGCODE"].Value;
			rsJoborderItemNew.Fields["FK_ITEM"].Value = rsJoborderItemOld.Fields["FK_ITEM"].Value;
			rsJoborderItemNew.Fields["QUANTITY"].Value = rsJoborderItemOld.Fields["QUANTITY"].Value;
			rsJoborderItemNew.Fields["LENGTH"].Value = rsJoborderItemOld.Fields["LENGTH"].Value;
			rsJoborderItemNew.Fields["WIDTH"].Value = rsJoborderItemOld.Fields["WIDTH"].Value;
			rsJoborderItemNew.Fields["CAMPARAMETER"].Value = rsJoborderItemOld.Fields["CAMPARAMETER"].Value;
			rsJoborderItemNew.Fields["MACHINENAMECAM"].Value = rsJoborderItemOld.Fields["MACHINENAMECAM"].Value;
			rsJoborderItemNew.Fields["DIM_W"].Value = rsJoborderItemOld.Fields["DIM_W"].Value;

			rsJoborderItemNew.Update();



			rsJoborderItemOld.Delete();
			//	rsJoborderItemOld.Update();








		}

	}
}