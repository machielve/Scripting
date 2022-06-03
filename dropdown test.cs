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
	
	private static DialogResult ShowInputDialog(ref decimal input1, ref DataTable dtTest, ref String input2)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Hastings";

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

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtTest;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
		inputBox.Controls.Add(combo1);


		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		input2 = combo1.SelectedValue.ToString();

		return result;
	}

	

	public void Execute()
	{
		decimal input1 = 1;
		string input2 = "";

		ScriptRecordset rsTest = this.GetRecordset("R_ITEM", "CODE, DESCRIPTION, DRAWINGNUMBER", string.Format("UNMARKETABLE = '{0}'", false), "DESCRIPTION");
		rsTest.MoveFirst();

		DataTable dtTest = rsTest.DataTable;

		DataColumn extracolumn = new DataColumn();
		extracolumn.DataType = System.Type.GetType("System.String");
		extracolumn.ColumnName = "TOTAAL";
		extracolumn.Expression = "(CODE)+(' - ')+(DESCRIPTION)+(' - ')+(DRAWINGNUMBER)";

		dtTest.Columns.Add(extracolumn);
		
		ShowInputDialog(ref input1, ref dtTest, ref input2);

		MessageBox.Show(input2);
	}


}
