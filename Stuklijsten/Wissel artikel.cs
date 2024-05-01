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
	private static DialogResult ShowInputDialog(ref string input1, ref DataTable dtItems)
	{
		System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
		customCulture.NumberFormat.NumberDecimalSeparator = ",";

		System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

		System.Drawing.Size size = new System.Drawing.Size(300, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Castillon";

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
		groepprijs.Text = "Nieuw artikel:";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtItems;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 25);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		groepprijs.Controls.Add(combo1);

		inputBox.Controls.Add(groepprijs);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();
		input1 = combo1.SelectedValue.ToString();

		return result;
	}

	public void Execute()
	{
		ScriptRecordset rsItemList = this.GetRecordset("R_ITEM", "CODE, DESCRIPTION, DRAWINGNUMBER", string.Format("UNMARKETABLE = '{0}'", false), "DESCRIPTION");
		rsItemList.MoveFirst();

		DataTable dtItems = rsItemList.DataTable;

		DataColumn extracolumn = new DataColumn();
		extracolumn.DataType = System.Type.GetType("System.String");
		extracolumn.ColumnName = "TOTAAL";
		extracolumn.Expression = "(CODE)+(' - ')+(DESCRIPTION)+(' - ')+(DRAWINGNUMBER)";

		dtItems.Columns.Add(extracolumn);



		string input1 = "";
		ShowInputDialog(ref input1, ref dtItems);

		ScriptRecordset rsItemNew = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", input1), "");
		rsItemNew.MoveFirst();

		int ItemID = Convert.ToInt32(rsItemNew.Fields["PK_R_ITEM"].Value.ToString());
		string Name = rsItemNew.Fields["DESCRIPTION"].Value.ToString();
		decimal maxlang = Convert.ToDecimal(rsItemNew.Fields["TRADELENGTH"].Value.ToString());
		decimal maxbreed = Convert.ToDecimal(rsItemNew.Fields["TRADEWIDTH"].Value.ToString());


		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsItemRow = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM = " + (int)record.GetPrimaryKeyValue(), "");
			rsItemRow.MoveFirst();
			rsItemRow.UseDataChanges = true;

			// oude info
			decimal aantal = Convert.ToDecimal(rsItemRow.Fields["QUANTITY"].Value.ToString());
			decimal Lengte = Convert.ToDecimal(rsItemRow.Fields["LENGTH"].Value.ToString());
			decimal Breedte = Convert.ToDecimal(rsItemRow.Fields["WIDTH"].Value.ToString());
			
			if( Lengte > maxlang)
				Lengte = maxlang;
			
			if( Breedte > maxbreed)
				Breedte = maxbreed;

			if (Lengte == 0 && maxlang > 0)
				Lengte = maxlang;
			
			if (Breedte == 0 && maxbreed > 0)
				Breedte = maxbreed;

			// nieuwe info
			rsItemRow.Fields["FK_ITEM"].Value = ItemID;
			rsItemRow.Fields["DESCRIPTION"].Value = Name;
			rsItemRow.Fields["QUANTITY"].Value = aantal;
			rsItemRow.Fields["LENGTH"].Value = Lengte;
			rsItemRow.Fields["WIDTH"].Value = Breedte;



			rsItemRow.Update(null, null);

		}


	}
}