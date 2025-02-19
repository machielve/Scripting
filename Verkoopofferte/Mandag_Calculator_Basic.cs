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
	
	Mandag calculator, het  programma om het benodigde aantal mandagen voor de montage te berekenen
	Uit te voeren vanuit een offerte met de status nieuw
	Geschreven door: Machiel R. van Emden feb-2025

	*/

	private static DialogResult ShowInputDialog(ref decimal input1,
													ref decimal input2,
													ref decimal input3,
													ref decimal input4,
													ref decimal input5,
													ref decimal fixed1,
													ref decimal fixed2,
													ref decimal fixed3,
													ref decimal fixed4,
													ref decimal fixed5)
	{

		System.Drawing.Size size = new System.Drawing.Size(350, 300);
		Form inputBox = new Form();

		int c0 = 5;             //labelin
		int c1 = c0 + 100;      //input
		int c2 = c1 + 80;       //input unit
		int c3 = c2 + 50;       //delen of vermenigvuldigen
		int c4 = c3 + 10;       //fixed
		int c5 = c4 + 40;       //=...

		int r1 = 55;
		int r2 = r1 + 35;
		int r3 = r2 + 35;
		int r4 = r3 + 35;
		int r5 = r4 + 35;

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.Icon = new System.Drawing.Icon(@"W:\Machiel\Ridder\Scripting\icons\werkman.ico");
		inputBox.ClientSize = size;
		inputBox.Text = "Mandag Calculatort";

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



		// kolom 1 - omschrijvingen

		System.Windows.Forms.Label labelin1 = new Label();
		labelin1.Size = new System.Drawing.Size(100, 25);
		labelin1.Location = new System.Drawing.Point(c0, r1 + 2);
		labelin1.Text = "Oppervlakte vloer";
		inputBox.Controls.Add(labelin1);

		System.Windows.Forms.Label labelin2 = new Label();
		labelin2.Size = new System.Drawing.Size(100, 25);
		labelin2.Location = new System.Drawing.Point(c0, r2 + 2);
		labelin2.Text = "Aantal trappen";
		inputBox.Controls.Add(labelin2);

		System.Windows.Forms.Label labelin3 = new Label();
		labelin3.Size = new System.Drawing.Size(100, 25);
		labelin3.Location = new System.Drawing.Point(c0, r3 + 2);
		labelin3.Text = "Aantal POP";
		inputBox.Controls.Add(labelin3);

		System.Windows.Forms.Label labelin4 = new Label();
		labelin4.Size = new System.Drawing.Size(100, 25);
		labelin4.Location = new System.Drawing.Point(c0, r4 + 2);
		labelin4.Text = "Aantal m leuning";
		inputBox.Controls.Add(labelin4);

		System.Windows.Forms.Label labelin5 = new Label();
		labelin5.Size = new System.Drawing.Size(100, 25);
		labelin5.Location = new System.Drawing.Point(c0, r5 + 2);
		labelin5.Text = "Aantal m hekwerk";
		inputBox.Controls.Add(labelin5);


		// kolom 2 - input velden

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(75, 25);
		textBox1.Location = new System.Drawing.Point(c1, r1);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 1500000;
		textBox1.DecimalPlaces = 2;
		inputBox.Controls.Add(textBox1);

		System.Windows.Forms.NumericUpDown textBox2 = new NumericUpDown();
		textBox2.Size = new System.Drawing.Size(75, 25);
		textBox2.Location = new System.Drawing.Point(c1, r2);
		textBox2.Value = input2;
		textBox2.Minimum = 0;
		textBox2.Maximum = 1500000;
		textBox2.DecimalPlaces = 0;
		inputBox.Controls.Add(textBox2);

		System.Windows.Forms.NumericUpDown textBox3 = new NumericUpDown();
		textBox3.Size = new System.Drawing.Size(75, 25);
		textBox3.Location = new System.Drawing.Point(c1, r3);
		textBox3.Value = input3;
		textBox3.Minimum = 0;
		textBox3.Maximum = 1500000;
		textBox3.DecimalPlaces = 0;
		inputBox.Controls.Add(textBox3);

		System.Windows.Forms.NumericUpDown textBox4 = new NumericUpDown();
		textBox4.Size = new System.Drawing.Size(75, 25);
		textBox4.Location = new System.Drawing.Point(c1, r4);
		textBox4.Value = input4;
		textBox4.Minimum = 0;
		textBox4.Maximum = 1500000;
		textBox4.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox4);

		System.Windows.Forms.NumericUpDown textBox5 = new NumericUpDown();
		textBox5.Size = new System.Drawing.Size(75, 25);
		textBox5.Location = new System.Drawing.Point(c1, r5);
		textBox5.Value = input5;
		textBox5.Minimum = 0;
		textBox5.Maximum = 1500000;
		textBox5.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox5);



		// kolom 3 - eenheid van input

		System.Windows.Forms.Label labelunit1 = new Label();
		labelunit1.Size = new System.Drawing.Size(50, 25);
		labelunit1.Location = new System.Drawing.Point(c2, r1 + 2);
		labelunit1.Text = "m²";
		inputBox.Controls.Add(labelunit1);

		System.Windows.Forms.Label labelunit2 = new Label();
		labelunit2.Size = new System.Drawing.Size(50, 25);
		labelunit2.Location = new System.Drawing.Point(c2, r2 + 2);
		labelunit2.Text = "stuks";
		inputBox.Controls.Add(labelunit2);

		System.Windows.Forms.Label labelunit3 = new Label();
		labelunit3.Size = new System.Drawing.Size(50, 25);
		labelunit3.Location = new System.Drawing.Point(c2, r3 + 2);
		labelunit3.Text = "stuks";
		inputBox.Controls.Add(labelunit3);

		System.Windows.Forms.Label labelunit4 = new Label();
		labelunit4.Size = new System.Drawing.Size(50, 25);
		labelunit4.Location = new System.Drawing.Point(c2, r4 + 2);
		labelunit4.Text = "m";
		inputBox.Controls.Add(labelunit4);

		System.Windows.Forms.Label labelunit5 = new Label();
		labelunit5.Size = new System.Drawing.Size(50, 25);
		labelunit5.Location = new System.Drawing.Point(c2, r5 + 2);
		labelunit5.Text = "m";
		inputBox.Controls.Add(labelunit5);


		// kolom 4 - delen of vermenigvuldigen

		System.Windows.Forms.Label labelexp1 = new Label();
		labelexp1.Size = new System.Drawing.Size(10, 25);
		labelexp1.Location = new System.Drawing.Point(c3, r1 + 2);
		labelexp1.Text = @"/";
		inputBox.Controls.Add(labelexp1);

		System.Windows.Forms.Label labelexp2 = new Label();
		labelexp2.Size = new System.Drawing.Size(10, 25);
		labelexp2.Location = new System.Drawing.Point(c3, r2 + 2);
		labelexp2.Text = @"*";
		inputBox.Controls.Add(labelexp2);

		System.Windows.Forms.Label labelexp3 = new Label();
		labelexp3.Size = new System.Drawing.Size(10, 25);
		labelexp3.Location = new System.Drawing.Point(c3, r3 + 2);
		labelexp3.Text = @"*";
		inputBox.Controls.Add(labelexp3);

		System.Windows.Forms.Label labelexp4 = new Label();
		labelexp4.Size = new System.Drawing.Size(10, 25);
		labelexp4.Location = new System.Drawing.Point(c3, r4 + 2);
		labelexp4.Text = @"/";
		inputBox.Controls.Add(labelexp4);

		System.Windows.Forms.Label labelexp5 = new Label();
		labelexp5.Size = new System.Drawing.Size(10, 25);
		labelexp5.Location = new System.Drawing.Point(c3, r5 + 2);
		labelexp5.Text = @"/";
		inputBox.Controls.Add(labelexp5);


		// kolom 5 - vaste rekenwaardes 

		System.Windows.Forms.NumericUpDown textBox21 = new NumericUpDown();
		textBox21.Size = new System.Drawing.Size(40, 25);
		textBox21.Location = new System.Drawing.Point(c4, r1);
		textBox21.Value = fixed1;
		textBox21.Minimum = 0;
		textBox21.Maximum = 1500000;
		textBox21.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox21);
		textBox21.Controls.RemoveAt(0);
		textBox21.BackColor = Color.LightGray;

		System.Windows.Forms.NumericUpDown textBox22 = new NumericUpDown();
		textBox22.Size = new System.Drawing.Size(40, 25);
		textBox22.Location = new System.Drawing.Point(c4, r2);
		textBox22.Value = fixed2;
		textBox22.Minimum = 0;
		textBox22.Maximum = 1500000;
		textBox22.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox22);
		textBox22.Controls.RemoveAt(0);
		textBox22.BackColor = Color.LightGray;

		System.Windows.Forms.NumericUpDown textBox23 = new NumericUpDown();
		textBox23.Size = new System.Drawing.Size(40, 25);
		textBox23.Location = new System.Drawing.Point(c4, r3);
		textBox23.Value = fixed3;
		textBox23.Minimum = 0;
		textBox23.Maximum = 1500000;
		textBox23.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox23);
		textBox23.Controls.RemoveAt(0);
		textBox23.BackColor = Color.LightGray;

		System.Windows.Forms.NumericUpDown textBox24 = new NumericUpDown();
		textBox24.Size = new System.Drawing.Size(40, 25);
		textBox24.Location = new System.Drawing.Point(c4, r4);
		textBox24.Value = fixed4;
		textBox24.Minimum = 0;
		textBox24.Maximum = 1500000;
		textBox24.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox24);
		textBox24.Controls.RemoveAt(0);
		textBox24.BackColor = Color.LightGray;

		System.Windows.Forms.NumericUpDown textBox25 = new NumericUpDown();
		textBox25.Size = new System.Drawing.Size(40, 25);
		textBox25.Location = new System.Drawing.Point(c4, r5);
		textBox25.Value = fixed5;
		textBox25.Minimum = 0;
		textBox25.Maximum = 1500000;
		textBox25.DecimalPlaces = 1;
		inputBox.Controls.Add(textBox25);
		textBox25.Controls.RemoveAt(0);
		textBox25.BackColor = Color.LightGray;


		// kolom 6 - delen of vermenigvuldigen

		System.Windows.Forms.Label labeleq1 = new Label();
		labeleq1.Size = new System.Drawing.Size(50, 25);
		labeleq1.Location = new System.Drawing.Point(c5, r1 + 2);
		labeleq1.Text = @"= ....";
		inputBox.Controls.Add(labeleq1);

		System.Windows.Forms.Label labeleq2 = new Label();
		labeleq2.Size = new System.Drawing.Size(50, 25);
		labeleq2.Location = new System.Drawing.Point(c5, r2 + 2);
		labeleq2.Text = @"= ....";
		inputBox.Controls.Add(labeleq2);

		System.Windows.Forms.Label labeleq3 = new Label();
		labeleq3.Size = new System.Drawing.Size(50, 25);
		labeleq3.Location = new System.Drawing.Point(c5, r3 + 2);
		labeleq3.Text = @"= ....";
		inputBox.Controls.Add(labeleq3);

		System.Windows.Forms.Label labeleq4 = new Label();
		labeleq4.Size = new System.Drawing.Size(50, 25);
		labeleq4.Location = new System.Drawing.Point(c5, r4 + 2);
		labeleq4.Text = @"= ....";
		inputBox.Controls.Add(labeleq4);

		System.Windows.Forms.Label labeleq5 = new Label();
		labeleq5.Size = new System.Drawing.Size(50, 25);
		labeleq5.Location = new System.Drawing.Point(c5, r5 + 2);
		labeleq5.Text = @"= ....";
		inputBox.Controls.Add(labeleq5);


		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		input1 = textBox1.Value;
		input2 = textBox2.Value;
		input3 = textBox3.Value;
		input4 = textBox4.Value;
		input5 = textBox5.Value;

		fixed1 = textBox21.Value;
		fixed2 = textBox22.Value;
		fixed3 = textBox23.Value;
		fixed4 = textBox24.Value;
		fixed5 = textBox25.Value;

		return result;


	}


	public void Execute()
	{
		decimal input1 = 10;
		decimal input2 = 0;
		decimal input3 = 0;
		decimal input4 = 0;
		decimal input5 = 0;

		decimal fixed1 = 20;
		decimal fixed2 = 0.5m;
		decimal fixed3 = 1;
		decimal fixed4 = 15;
		decimal fixed5 = 20;

		DialogResult result = ShowInputDialog(ref input1, ref input2, ref input3, ref input4, ref input5, ref fixed1, ref fixed2, ref fixed3, ref fixed4, ref fixed5);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Mandag calculator afgebroken");
			return;
		}

		decimal output1 = (input1 / fixed1) + (input2 * fixed2) + (input3 * fixed3) + (input4 / fixed4) + (input5 / fixed5);

		decimal output2 = Math.Ceiling(output1);

		double output3 = Convert.ToDouble(output2);

		double output4 = 0;

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;

		foreach (IRecord record in records)
		{
			ScriptRecordset rsOffer = this.GetRecordset("R_OFFER", "", "PK_R_OFFER = " + (int)record.GetPrimaryKeyValue(), "");
			rsOffer.MoveFirst();
			rsOffer.UseDataChanges = true;

			if (output3 < 2)
			{
				output4 = 2;
			}

			else
			{
				output4 = output3;
			}

			double montagetijd = output4 * 24 * 60 * 60 * 10000000;


			rsOffer.Fields["GESCHATMONTAGETIJD"].Value = montagetijd;

			rsOffer.Update();

		}

	}

	// M.R.v.E - 2025

}
