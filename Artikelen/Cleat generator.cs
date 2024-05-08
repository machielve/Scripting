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
	
	Cleat aanmaken, het  programma om te checken of de benodigde cleat al bestaat en zonodig een nieuwe genereerd.
	Uit te voeren vanuit de artikelen lijst
	Geschreven door: Machiel R. van Emden oktober-2022

	*/

	private static DialogResult ShowInputDialog(ref decimal input1, ref decimal input2, ref decimal input3, ref decimal input4,
												ref bool rb10, ref bool rb11, ref bool rb12, ref bool rb13,
												ref bool rb20, ref bool rb21, ref bool rb22,
												ref bool cb1, ref bool cb2)
	{

		System.Drawing.Size size = new System.Drawing.Size(500, 750);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Agincourt";

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


		//groep afmeting
		GroupBox groephoog = new GroupBox();
		groephoog.Size = new System.Drawing.Size(180, 150);
		groephoog.Location = new System.Drawing.Point(10, 200);
		groephoog.Text = "Cleat afmetingen";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(50, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 4000;
		groephoog.Controls.Add(textBox1);

		System.Windows.Forms.Label label1 = new Label();
		label1.Size = new System.Drawing.Size(40, 25);
		label1.Location = new System.Drawing.Point(5, 30);
		label1.Text = "L1=";
		groephoog.Controls.Add(label1);

		System.Windows.Forms.NumericUpDown textBox2 = new NumericUpDown();
		textBox2.Size = new System.Drawing.Size(100, 25);
		textBox2.Location = new System.Drawing.Point(50, 50);
		textBox2.Value = input2;
		textBox2.Minimum = 0;
		textBox2.Maximum = 4000;
		groephoog.Controls.Add(textBox2);

		System.Windows.Forms.Label label2 = new Label();
		label2.Size = new System.Drawing.Size(40, 25);
		label2.Location = new System.Drawing.Point(5, 55);
		label2.Text = "L2=";
		groephoog.Controls.Add(label2);

		System.Windows.Forms.NumericUpDown textBox3 = new NumericUpDown();
		textBox3.Size = new System.Drawing.Size(100, 25);
		textBox3.Location = new System.Drawing.Point(50, 75);
		textBox3.Value = input3;
		textBox3.Minimum = 0;
		textBox3.Maximum = 4000;
		groephoog.Controls.Add(textBox3);

		System.Windows.Forms.Label label3 = new Label();
		label3.Size = new System.Drawing.Size(40, 25);
		label3.Location = new System.Drawing.Point(5, 80);
		label3.Text = "t=";
		groephoog.Controls.Add(label3);

		System.Windows.Forms.NumericUpDown textBox4 = new NumericUpDown();
		textBox4.Size = new System.Drawing.Size(100, 25);
		textBox4.Location = new System.Drawing.Point(50, 100);
		textBox4.Value = input4;
		textBox4.Minimum = 0;
		textBox4.Maximum = 4000;
		groephoog.Controls.Add(textBox4);

		System.Windows.Forms.Label label4 = new Label();
		label4.Size = new System.Drawing.Size(40, 25);
		label4.Location = new System.Drawing.Point(5, 105);
		label4.Text = "H=";
		groephoog.Controls.Add(label4);

		inputBox.Controls.Add(groephoog);



		//groep type
		GroupBox groeptype = new GroupBox();
		groeptype.Size = new System.Drawing.Size(180, 140);
		groeptype.Location = new System.Drawing.Point(10, 50);
		groeptype.Text = "Cleat type";

		System.Windows.Forms.RadioButton rbutton10 = new RadioButton();
		rbutton10.Size = new System.Drawing.Size(150, 25);
		rbutton10.Location = new System.Drawing.Point(10, 25);
		rbutton10.Checked = true;
		rbutton10.Text = "Joist cleat";
		groeptype.Controls.Add(rbutton10);

		System.Windows.Forms.RadioButton rbutton11 = new RadioButton();
		rbutton11.Size = new System.Drawing.Size(150, 25);
		rbutton11.Location = new System.Drawing.Point(10, 50);
		rbutton11.Checked = false;
		rbutton11.Text = "Beam cleat";
		groeptype.Controls.Add(rbutton11);

		System.Windows.Forms.RadioButton rbutton12 = new RadioButton();
		rbutton12.Size = new System.Drawing.Size(150, 25);
		rbutton12.Location = new System.Drawing.Point(10, 75);
		rbutton12.Checked = false;
		rbutton12.Text = "Column cleat";
		groeptype.Controls.Add(rbutton12);

		System.Windows.Forms.RadioButton rbutton13 = new RadioButton();
		rbutton13.Size = new System.Drawing.Size(150, 25);
		rbutton13.Location = new System.Drawing.Point(10, 100);
		rbutton13.Checked = false;
		rbutton13.Text = "Sigma cleat";
		groeptype.Controls.Add(rbutton13);







		inputBox.Controls.Add(groeptype);


		//groep materiaal
		GroupBox groepmateriaal = new GroupBox();
		groepmateriaal.Size = new System.Drawing.Size(180, 125);
		groepmateriaal.Location = new System.Drawing.Point(10, 375);
		groepmateriaal.Text = "Staal kwaliteit";

		System.Windows.Forms.RadioButton rbutton20 = new RadioButton();
		rbutton20.Size = new System.Drawing.Size(75, 25);
		rbutton20.Location = new System.Drawing.Point(10, 25);
		rbutton20.Checked = true;
		rbutton20.Text = "S235JR";
		groepmateriaal.Controls.Add(rbutton20);

		System.Windows.Forms.RadioButton rbutton21 = new RadioButton();
		rbutton21.Size = new System.Drawing.Size(75, 25);
		rbutton21.Location = new System.Drawing.Point(10, 50);
		rbutton21.Checked = false;
		rbutton21.Text = "S275JR";
		groepmateriaal.Controls.Add(rbutton21);

		System.Windows.Forms.RadioButton rbutton22 = new RadioButton();
		rbutton22.Size = new System.Drawing.Size(75, 25);
		rbutton22.Location = new System.Drawing.Point(10, 75);
		rbutton22.Checked = false;
		rbutton22.Text = "S355JR";
		groepmateriaal.Controls.Add(rbutton22);

		inputBox.Controls.Add(groepmateriaal);


		//groep modificator
		GroupBox groepMod = new GroupBox();
		groepMod.Size = new System.Drawing.Size(180, 125);
		groepMod.Location = new System.Drawing.Point(10, 525);
		groepMod.Text = "Extra naam deel";

		System.Windows.Forms.CheckBox cbutton1 = new CheckBox();
		cbutton1.Size = new System.Drawing.Size(75, 25);
		cbutton1.Location = new System.Drawing.Point(10, 25);
		cbutton1.Checked = false;
		cbutton1.Text = "Multicleat";
		groepMod.Controls.Add(cbutton1);

		System.Windows.Forms.CheckBox cbutton2 = new CheckBox();
		cbutton2.Size = new System.Drawing.Size(75, 25);
		cbutton2.Location = new System.Drawing.Point(10, 50);
		cbutton2.Checked = false;
		cbutton2.Text = "Braced";
		groepMod.Controls.Add(cbutton2);

		inputBox.Controls.Add(groepMod);


		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;


		DialogResult result = inputBox.ShowDialog();
		input1 = textBox1.Value;
		input2 = textBox2.Value;
		input3 = textBox3.Value;
		input4 = textBox4.Value;

		rb10 = rbutton10.Checked;
		rb11 = rbutton11.Checked;
		rb12 = rbutton12.Checked;
		rb13 = rbutton13.Checked;

		rb20 = rbutton20.Checked;
		rb21 = rbutton21.Checked;
		rb22 = rbutton22.Checked;

		cb1 = cbutton1.Checked;
		cb2 = cbutton2.Checked;


		return result;
	}

	public void Execute()
	{
		decimal input1 = 60;
		decimal input2 = 60;
		decimal input3 = 4;
		decimal input4 = 100;

		bool rb10 = true;
		bool rb11 = false;
		bool rb12 = false;
		bool rb13 = false;

		bool rb20 = true;
		bool rb21 = false;
		bool rb22 = false;

		bool cb1 = false;
		bool cb2 = false;


		DialogResult result = ShowInputDialog(ref input1, ref input2, ref input3, ref input4,
							ref rb10, ref rb11, ref rb12, ref rb13,
							ref rb20, ref rb21, ref rb22,
							ref cb1, ref cb2);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Cleat generator afgebroken");
			return;
		}


		string sinput1 = input1.ToString();
		string sinput2 = input2.ToString();
		string sinput3 = input3.ToString();
		string sinput4 = input4.ToString();

		string voorzet = "";
		string mod1 = "";
		string mod2 = "";
		string materiaal = "";

		if (rb10 == true) voorzet = "Joist cleat ";
		if (rb11 == true) voorzet = "Beam cleat ";
		if (rb12 == true) voorzet = "Column cleat ";
		if (rb13 == true) voorzet = "Sigma cleat ";


		if (cb1 == true) mod1 = "multi ";
		if (cb2 == true) mod2 = "braced ";


		if (rb20 == true) materiaal = " S235JR";
		if (rb21 == true) materiaal = " S275JR";
		if (rb22 == true) materiaal = " S355JR";



		string maat1 = sinput1 + "x" + sinput2 + "x" + sinput3 + " H=" + sinput4;
		string maat2 = sinput2 + "x" + sinput1 + "x" + sinput3 + " H=" + sinput4;

		string naam1 = voorzet + mod1 + mod2 + maat1 + materiaal;
		string naam2 = voorzet + mod1 + mod2 + maat2 + materiaal;

		decimal oppervlak;
		decimal volume;

		if (rb12 == true)
		{
			oppervlak = (input1 + input2 + input1) * input4 * 2;
			volume = (input1 + input2 + input1) * input4 * input3;
		}

		else
		{
			oppervlak = (input1 + input2) * input4 * 2;
			volume = (input1 + input2) * input4 * input3;
		}



		decimal gewicht = volume / 1000000000 * 7850;
		decimal spuitvlak = oppervlak / 1000000;

		int AGroep = 116;
		int AEenheid = 33;
		int zaagcode = 1;
		int prijsupdate = 2;
		int insjabloon = 2;
		int versjabloon = 2;
		int rTraject = 3;

		int BCBid = 796;
		int Tailorid = 732;
		int Laserid = 764;
		int LaserMaxid = 837;
		int SiemSid = 836;



		ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("DESCRIPTION = '{0}'", naam1), "");
		rsItem.MoveFirst();

		ScriptRecordset rsItem2 = this.GetRecordset("R_ITEM", "", string.Format("DESCRIPTION = '{0}'", naam2), "");
		rsItem2.MoveFirst();


		if (rsItem.RecordCount > 0 || rsItem2.RecordCount > 0)
		{
			if (rsItem2.RecordCount > 0)
			{
				string ItemCode2 = rsItem2.Fields["CODE"].Value.ToString();
				MessageBox.Show("Artikel bestaat al onder: " + ItemCode2);
			}
			else
			{
				string ItemCode = rsItem.Fields["CODE"].Value.ToString();
				MessageBox.Show("Artikel bestaat al onder: " + ItemCode);
			}
		}

		else
		{
			rsItem.AddNew();
			rsItem.Fields["DESCRIPTION"].Value = naam1;
			rsItem.Fields["FK_ITEMGROUP"].Value = AGroep;
			rsItem.Fields["FK_ITEMUNIT"].Value = AEenheid;
			rsItem.Fields["DEFAULTSAWINGCODE"].Value = zaagcode;
			rsItem.Fields["FK_ITEMPURCHASEPRICETEMPLATEGROUP"].Value = insjabloon;
			rsItem.Fields["FK_ITEMSALESPRICETEMPLATEGROUP"].Value = versjabloon;
			rsItem.Fields["WEIGHT"].Value = gewicht;
			rsItem.Fields["REGISTRATIONPATH"].Value = rTraject;
			rsItem.Fields["PAINTAREA"].Value = spuitvlak;

			rsItem.Update();

			string NCode = rsItem.Fields["CODE"].Value.ToString();
			MessageBox.Show("Nieuw artikel = " + NCode + " - " + naam1);


			ScriptRecordset rsItemSup = this.GetRecordset("R_ITEMSUPPLIER", "", "PK_R_ITEMSUPPLIER = -1", "");
			rsItemSup.MoveFirst();

			// Hoofdleverancier als eerst
			rsItemSup.AddNew();
			rsItemSup.Fields["FK_RELATION"].Value = LaserMaxid;
			rsItemSup.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
			rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam1;
			rsItemSup.Fields["ITEMTYPE"].Value = 1;
			rsItemSup.Update();

			// extra leveranciers erna
			rsItemSup.AddNew();
			rsItemSup.Fields["FK_RELATION"].Value = Laserid;
			rsItemSup.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
			rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam1;
			rsItemSup.Fields["ITEMTYPE"].Value = 6;
			rsItemSup.Update();

			rsItemSup.AddNew();
			rsItemSup.Fields["FK_RELATION"].Value = Tailorid;
			rsItemSup.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
			rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam1;
			rsItemSup.Fields["ITEMTYPE"].Value = 6;
			rsItemSup.Update();

			rsItemSup.AddNew();
			rsItemSup.Fields["FK_RELATION"].Value = SiemSid;
			rsItemSup.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
			rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam1;
			rsItemSup.Fields["ITEMTYPE"].Value = 5;
			rsItemSup.Update();
			
			/*

			rsItemSup.AddNew();
			rsItemSup.Fields["FK_RELATION"].Value = BCBid;
			rsItemSup.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
			rsItemSup.Fields["PURCHASEDESCRIPTION"].Value = naam1;
			rsItemSup.Fields["ITEMTYPE"].Value = 2;
			rsItemSup.Update();

*/

		}
	}
}

// M.R.v.E - 2022

