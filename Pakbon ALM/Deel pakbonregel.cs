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
		inputBox.Text = "Word niet gelezen";

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
		string input = "Aantal regels";
		DialogResult result = ShowInputDialog(ref input);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Deel pakbonregel afgebroken");
			return;
		}
		
		int input1 = Convert.ToInt32(input);

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("U_PACKLISTDETAILITEM", "", "PK_U_PACKLISTDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
			rsItem.UseDataChanges = true;

			int aantal = Convert.ToInt32(rsItem.Fields["QUANTITY"].Value.ToString());

			if (input1 > aantal)
			{
				MessageBox.Show("Aantal regels hoger dan aantal artikelen");
				continue;
			}

			int nieuwaantal = (int)(aantal / input1);

			rsItem.Fields["QUANTITY"].Value = nieuwaantal;


			for (int i = 0; i < input1-1; i++)
			{
			//	MessageBox.Show(nieuwaantal.ToString());
				
				ScriptRecordset rsPakRegel = this.GetRecordset("U_PACKLISTDETAILITEM", "", "PK_U_PACKLISTDETAILITEM= -1", "");
				rsPakRegel.MoveFirst();
				rsPakRegel.AddNew();
				rsPakRegel.Fields["FK_BONREGELART"].Value = rsItem.Fields["FK_BONREGELART"].Value;
				rsPakRegel.Fields["QUANTITY"].Value = nieuwaantal;
				rsPakRegel.Fields["FK_PACKLIST"].Value = rsItem.Fields["FK_PACKLIST"].Value;
				rsPakRegel.Update();
				
			}
			
			
		

			rsItem.Update(null, null);


		}


	}
}