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
		string input = "Afwerking";
		DialogResult result = ShowInputDialog(ref input);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Afwerking invullen afgebroken");
			return;
		}

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItem = this.GetRecordset("R_JOBORDERDETAILITEM", "AFWERKING", "PK_R_JOBORDERDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItem.MoveFirst();
		//	rsItem.UseDataChanges = true;

			rsItem.Fields["AFWERKING"].Value = input;

			rsItem.Update();


		}


	}
}