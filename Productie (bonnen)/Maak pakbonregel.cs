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
		inputBox.Text = "Pakbon nummer";

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
			ScriptRecordset rsItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();

			decimal aantal = Convert.ToDecimal(rsItem.Fields["QUANTITY"].Value.ToString());
			string bonS =  rsItem.Fields["FK_JOBORDER"].Value.ToString();
			int bon = Convert.ToInt32(bonS);			
			string bonregelS = rsItem.Fields["PK_R_JOBORDERDETAILITEM"].Value.ToString();
			int bonregel = Convert.ToInt32(bonregelS);


			// controle of de regel overgeslagen moet worden
			if (rsItem.Fields["DIRECTELEVERING"].Value.ToString() == "False")
			{
				MessageBox.Show("Pakbonregel uitgevinkt. Regel is overgeslagen.");
				continue;
			}	
			
				
			
		

			
			// achterhalen van pakbon id nummer
			int pakboner = 0;

			ScriptRecordset rsPakbon = this.GetRecordset("U_PACKLIST", "", "FK_JOBORDER= " + (int)bon, "PACKLISTNUMBER");
			rsPakbon.MoveFirst();

			while (rsPakbon.EOF == false)
			{ 
				if (rsPakbon.Fields["PACKLISTNUMBER"].Value.ToString() == input.ToString())
				{
					pakboner += Convert.ToInt32(rsPakbon.Fields["PK_U_PACKLIST"].Value.ToString());
					rsPakbon.MoveNext();
				}
				else rsPakbon.MoveNext();	
			}



			// uitrekenen hoeveel per regel al gebruikt is
			int totaal = 0;
			
			ScriptRecordset rsAanwezig = this.GetRecordset("U_PACKLISTDETAILITEM", "", "FK_BONREGELART = " + (int)bonregel, "");
			rsAanwezig.MoveFirst();

			while (rsAanwezig.EOF == false)
			{
				totaal += Convert.ToInt32(rsAanwezig.Fields["QUANTITY"].Value.ToString());
				rsAanwezig.MoveNext();
			}	

			// uitrekenen hoeveel op de regel toe gevoegd word en aanmaken van pakbon regel
			decimal mogelijk = aantal - totaal;

			if (mogelijk == 0)
			{
				MessageBox.Show("Maximaal aantal reeds op pakbon(nen)");
				continue;
			}
			else
			{
				ScriptRecordset rsPakRegel = this.GetRecordset("U_PACKLISTDETAILITEM", "", "PK_U_PACKLISTDETAILITEM= -1", "");
				rsPakRegel.MoveFirst();
				rsPakRegel.AddNew();

				rsPakRegel.Fields["FK_BONREGELART"].Value = rsItem.Fields["PK_R_JOBORDERDETAILITEM"].Value;
				rsPakRegel.Fields["QUANTITY"].Value = mogelijk;
				rsPakRegel.Fields["FK_PACKLIST"].Value = pakboner;
				rsPakRegel.Update();
			}
		}
	}
}