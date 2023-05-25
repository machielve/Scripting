using ADODB;
using System;
using System.IO;
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
	private static DialogResult Keuzeveld1(ref string gekozen, ref List<string> lijst)
	{
		System.Drawing.Size size = new System.Drawing.Size(500, 350);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Calculatiesheet keuze";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DataSource = lijst;
		combo1.Size = new System.Drawing.Size(325, 25);
		combo1.DropDownWidth = 450;
		combo1.Location = new System.Drawing.Point(5, 50);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 50);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "cancelButton";
		cancelButton.Size = new System.Drawing.Size(75, 23);
		cancelButton.Text = "&Cancel";
		cancelButton.Location = new System.Drawing.Point(size.Width - 80, size.Height - 50);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();
		gekozen = combo1.SelectedValue.ToString();

		return result;

	}
	
	public void Execute()
	{
		string gekozen = "";		
		
		DirectoryInfo folder = new DirectoryInfo(@"W:\Almacon Offertes\");
		if (folder.Exists) 
		{
			string[] files = Directory.GetFiles(@"W:\Almacon Offertes\", "Alldeck platform calculator *.xlsx");

			List<string> lijst = new List<string>(files);
			
			Keuzeveld1(ref gekozen, ref lijst);
			
			
			

			
			
		//	System.Diagnostics.Process.Start(gekozen);
			MessageBox.Show(gekozen);

		}
		else MessageBox.Show("Verkeerde map");

	}
}