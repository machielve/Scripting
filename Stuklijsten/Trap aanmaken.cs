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
	
	Trap aanmaken, het  programma om de benodigde onderdelen van een almacon trap in een stuklijst te syoppen
	Uit te voeren vanuit een stuklijst met de status engineering
	Geschreven door: Machiel R. van Emden mei-2022

	*/
	
	private static DialogResult ShowInputDialog(ref string input, ref decimal input1, ref string input2, ref bool rb10, ref bool rb11,
												ref bool rb0, ref bool rb1, ref bool rb2, ref bool rb3, ref bool rb4, ref bool rb5)
	{

		System.Drawing.Size size = new System.Drawing.Size(300, 700);
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


		//aantal trappen
		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(size.Width - 10, 25);
		textBox.Location = new System.Drawing.Point(5, 50);
		textBox.Text = input;
		inputBox.Controls.Add(textBox);

		//groep hoogte
		GroupBox groephoog = new GroupBox();
		groephoog.Size = new System.Drawing.Size(180, 60);
		groephoog.Location = new System.Drawing.Point(10, 75);
		groephoog.Text = "Trap hoogte";

		System.Windows.Forms.NumericUpDown textBox1 = new NumericUpDown();
		textBox1.Size = new System.Drawing.Size(100, 25);
		textBox1.Location = new System.Drawing.Point(5, 25);
		textBox1.Value = input1;
		textBox1.Minimum = 0;
		textBox1.Maximum = 4000;
		groephoog.Controls.Add(textBox1);

		inputBox.Controls.Add(groephoog);

		//groep breedte
		GroupBox groepbreed = new GroupBox();
		groepbreed.Size = new System.Drawing.Size(180, 60);
		groepbreed.Location = new System.Drawing.Point(10, 150);
		groepbreed.Text = "Trap breedte";

		Label breed = new Label();
		breed.Text = "Trap breedte";
		breed.Size = new System.Drawing.Size(75, 25);
		breed.Location = new System.Drawing.Point(5, 100);
		groepbreed.Controls.Add(breed);

		ComboBox BoxBreed = new ComboBox();
		BoxBreed.Size = new System.Drawing.Size(100, 25);
		BoxBreed.Location = new System.Drawing.Point(5, 25);
		BoxBreed.DropDownStyle = ComboBoxStyle.DropDown;
		BoxBreed.Items.Add("600");
		BoxBreed.Items.Add("800");
		BoxBreed.Items.Add("900");
		BoxBreed.Items.Add("1000");
		BoxBreed.Items.Add("1200");
		BoxBreed.SelectedIndex = 1;
		BoxBreed.Text = input2;
		groepbreed.Controls.Add(BoxBreed);

		inputBox.Controls.Add(groepbreed);


		//groep type
		GroupBox groepBoxType = new GroupBox();
		groepBoxType.Size = new System.Drawing.Size(180, 200);
		groepBoxType.Location = new System.Drawing.Point(10, 225);
		groepBoxType.Text = "Trap type";


		System.Windows.Forms.RadioButton rbutton0 = new RadioButton();
		rbutton0.Size = new System.Drawing.Size(75, 25);
		rbutton0.Location = new System.Drawing.Point(10, 25);
		rbutton0.Checked = rb0;
		rbutton0.Text = "Type 0";
		rbutton0.Checked = true;
		groepBoxType.Controls.Add(rbutton0);

		System.Windows.Forms.RadioButton rbutton1 = new RadioButton();
		rbutton1.Size = new System.Drawing.Size(75, 25);
		rbutton1.Location = new System.Drawing.Point(10, 50);
		rbutton1.Checked = rb1;
		rbutton1.Text = "Type 1";
		groepBoxType.Controls.Add(rbutton1);

		System.Windows.Forms.RadioButton rbutton2 = new RadioButton();
		rbutton2.Size = new System.Drawing.Size(75, 25);
		rbutton2.Location = new System.Drawing.Point(10, 75);
		rbutton2.Checked = rb2;
		rbutton2.Text = "Type 2";
		groepBoxType.Controls.Add(rbutton2);

		System.Windows.Forms.RadioButton rbutton3 = new RadioButton();
		rbutton3.Size = new System.Drawing.Size(75, 25);
		rbutton3.Location = new System.Drawing.Point(10, 100);
		rbutton3.Checked = rb3;
		rbutton3.Text = "Type 3";
		groepBoxType.Controls.Add(rbutton3);

		System.Windows.Forms.RadioButton rbutton4 = new RadioButton();
		rbutton4.Size = new System.Drawing.Size(75, 25);
		rbutton4.Location = new System.Drawing.Point(10, 125);
		rbutton4.Checked = rb4;
		rbutton4.Text = "Type 4";
		groepBoxType.Controls.Add(rbutton4);

		System.Windows.Forms.RadioButton rbutton5 = new RadioButton();
		rbutton5.Size = new System.Drawing.Size(75, 25);
		rbutton5.Location = new System.Drawing.Point(10, 150);
		rbutton5.Checked = rb5;
		rbutton5.Text = "Type 5";
		groepBoxType.Controls.Add(rbutton5);

		inputBox.Controls.Add(groepBoxType);

		//groephoek	
		GroupBox groepBoxHoek = new GroupBox();
		groepBoxHoek.Size = new System.Drawing.Size(180, 100);
		groepBoxHoek.Location = new System.Drawing.Point(10, 450);
		groepBoxHoek.Text = "Trap hoek";


		System.Windows.Forms.RadioButton rbutton10 = new RadioButton();
		rbutton10.Size = new System.Drawing.Size(75, 25);
		rbutton10.Location = new System.Drawing.Point(10, 25);
		rbutton10.Checked = rb10;
		rbutton10.Text = "42º";
		rbutton10.Checked = true;
		groepBoxHoek.Controls.Add(rbutton10);

		System.Windows.Forms.RadioButton rbutton11 = new RadioButton();
		rbutton11.Size = new System.Drawing.Size(75, 25);
		rbutton11.Location = new System.Drawing.Point(10, 50);
		rbutton11.Checked = rb11;
		rbutton11.Text = "37º";
		groepBoxHoek.Controls.Add(rbutton11);


		inputBox.Controls.Add(groepBoxHoek);


		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;


		DialogResult result = inputBox.ShowDialog();
		input = textBox.Text;
		input1 = textBox1.Value;
		input2 = BoxBreed.Text;

		rb10 = rbutton10.Checked;
		rb11 = rbutton11.Checked;

		rb0 = rbutton0.Checked;
		rb1 = rbutton1.Checked;
		rb2 = rbutton2.Checked;
		rb3 = rbutton3.Checked;
		rb4 = rbutton4.Checked;
		rb5 = rbutton5.Checked;

		return result;
	}

	public void Execute()
	{
		string input = "Aantal trappen";
		decimal input1 = 1;
		string input2 = "Trap breedte";
		bool rb10 = false;
		bool rb11 = false;


		bool rb0 = true;
		bool rb1 = false;
		bool rb2 = false;
		bool rb3 = false;
		bool rb4 = false;
		bool rb5 = false;


		ShowInputDialog(ref input, ref input1, ref input2, ref rb10, ref rb11, ref rb0, ref rb1, ref rb2, ref rb3, ref rb4, ref rb5);

		int hoek = 0;
		int type = 0;
		int ssm = 0;
		string model = "";
		string stuklijst = "";
		string tredecode = "";
		string supportcode = "";
		decimal treden = 0;
		decimal hoog = input1;
		decimal optreden42 = Math.Round(hoog / 210, 0);
		decimal optreden37 = Math.Round(hoog / 190, 0);
		decimal breed = Int32.Parse(input2);


		if (rb10 == true)
		{
			hoek = 42;
			treden = optreden42 - 1;

		}
		else if (rb11 == true)
		{
			hoek = 37;
			treden = optreden37 - 1;
		}
		else
		{
			hoek = 0;
			treden = 0;
		}

		if (rb0 == true)
		{
			type = 0;
			model = "A";
		}
		else if (rb1 == true)
		{
			type = 1;
			model = "A";
		}
		else if (rb2 == true)
		{
			type = 2;
			model = "B";
		}
		else if (rb3 == true)
		{
			type = 3;
			model = "B";
		}
		else if (rb4 == true)
		{
			type = 4;
			model = "B";
		}
		else if (rb5 == true)
		{
			type = 5;
			model = "B";
		}

		else
		{
			type = 10;
			model = "";
		}



		if (hoek == 37 && model == "A")
		{
			if (hoog >= 0 && hoog <= 190) { stuklijst = "S100021"; }
			else if (hoog <= 380) { stuklijst = "S100022"; }
			else if (hoog <= 570) { stuklijst = "S100023"; }
			else if (hoog <= 760) { stuklijst = "S100024"; }
			else if (hoog <= 950) { stuklijst = "S100025"; }
			else if (hoog <= 1140) { stuklijst = "S100026"; }
			else if (hoog <= 1330) { stuklijst = "S100027"; }
			else if (hoog <= 1520) { stuklijst = "S100028"; }
			else if (hoog <= 1710) { stuklijst = "S100029"; }
			else if (hoog <= 1900) { stuklijst = "S100030"; }
			else if (hoog <= 2090) { stuklijst = "S100031"; }
			else if (hoog <= 2280) { stuklijst = "S100032"; }
			else if (hoog <= 2470) { stuklijst = "S100033"; }
			else if (hoog <= 2660) { stuklijst = "S100034"; }
			else if (hoog <= 2850) { stuklijst = "S100035"; }
			else if (hoog <= 3040) { stuklijst = "S100036"; }
			else if (hoog <= 3230) { stuklijst = "S100037"; }
			else if (hoog <= 3420) { stuklijst = "S100038"; }
			else if (hoog <= 3610) { stuklijst = "S100039"; }
			else if (hoog <= 3800) { stuklijst = "S100040"; }
			else if (hoog <= 3990) { stuklijst = "S100041"; }
			else if (hoog <= 4001) { stuklijst = "S100042"; }
		}

		if (hoek == 37 && model == "B")
		{
			if (hoog >= 0 && hoog <= 190) { stuklijst = "S100043"; }
			else if (hoog <= 380) { stuklijst = "S100044"; }
			else if (hoog <= 570) { stuklijst = "S100045"; }
			else if (hoog <= 760) { stuklijst = "S100046"; }
			else if (hoog <= 950) { stuklijst = "S100047"; }
			else if (hoog <= 1140) { stuklijst = "S100048"; }
			else if (hoog <= 1330) { stuklijst = "S100049"; }
			else if (hoog <= 1520) { stuklijst = "S100050"; }
			else if (hoog <= 1710) { stuklijst = "S100051"; }
			else if (hoog <= 1900) { stuklijst = "S100052"; }
			else if (hoog <= 2090) { stuklijst = "S100053"; }
			else if (hoog <= 2280) { stuklijst = "S100054"; }
			else if (hoog <= 2470) { stuklijst = "S100055"; }
			else if (hoog <= 2660) { stuklijst = "S100056"; }
			else if (hoog <= 2850) { stuklijst = "S100057"; }
			else if (hoog <= 3040) { stuklijst = "S100058"; }
			else if (hoog <= 3230) { stuklijst = "S100059"; }
			else if (hoog <= 3420) { stuklijst = "S100060"; }
			else if (hoog <= 3610) { stuklijst = "S100061"; }
			else if (hoog <= 3800) { stuklijst = "S100062"; }
			else if (hoog <= 3990) { stuklijst = "S100063"; }
			else if (hoog <= 4001) { stuklijst = "S100064"; }
		}

		if (hoek == 42 && model == "A")
		{
			if (hoog >= 0 && hoog <= 210) { stuklijst = "S100065"; }
			else if (hoog <= 420) { stuklijst = "S100066"; }
			else if (hoog <= 630) { stuklijst = "S100067"; }
			else if (hoog <= 840) { stuklijst = "S100068"; }
			else if (hoog <= 1050) { stuklijst = "S100069"; }
			else if (hoog <= 1260) { stuklijst = "S100070"; }
			else if (hoog <= 1470) { stuklijst = "S100071"; }
			else if (hoog <= 1680) { stuklijst = "S100072"; }
			else if (hoog <= 1890) { stuklijst = "S100073"; }
			else if (hoog <= 2100) { stuklijst = "S100074"; }
			else if (hoog <= 2310) { stuklijst = "S100075"; }
			else if (hoog <= 2520) { stuklijst = "S100076"; }
			else if (hoog <= 2730) { stuklijst = "S100077"; }
			else if (hoog <= 2940) { stuklijst = "S100078"; }
			else if (hoog <= 3150) { stuklijst = "S100079"; }
			else if (hoog <= 3360) { stuklijst = "S100080"; }
			else if (hoog <= 3570) { stuklijst = "S100081"; }
			else if (hoog <= 3780) { stuklijst = "S100082"; }
			else if (hoog <= 3990) { stuklijst = "S100083"; }
			else if (hoog <= 4001) { stuklijst = "S100084"; }
		}

		if (hoek == 42 && model == "B")
		{
			if (hoog >= 0 && hoog <= 210) { stuklijst = "S100085"; }
			else if (hoog <= 420) { stuklijst = "S100086"; }
			else if (hoog <= 630) { stuklijst = "S100087"; }
			else if (hoog <= 840) { stuklijst = "S100088"; }
			else if (hoog <= 1050) { stuklijst = "S100089"; }
			else if (hoog <= 1260) { stuklijst = "S100090"; }
			else if (hoog <= 1470) { stuklijst = "S100091"; }
			else if (hoog <= 1680) { stuklijst = "S100092"; }
			else if (hoog <= 1890) { stuklijst = "S100093"; }
			else if (hoog <= 2100) { stuklijst = "S100094"; }
			else if (hoog <= 2310) { stuklijst = "S100095"; }
			else if (hoog <= 2520) { stuklijst = "S100096"; }
			else if (hoog <= 2730) { stuklijst = "S100097"; }
			else if (hoog <= 2940) { stuklijst = "S100098"; }
			else if (hoog <= 3150) { stuklijst = "S100099"; }
			else if (hoog <= 3360) { stuklijst = "S100100"; }
			else if (hoog <= 3570) { stuklijst = "S100101"; }
			else if (hoog <= 3780) { stuklijst = "S100102"; }
			else if (hoog <= 3990) { stuklijst = "S100103"; }
			else if (hoog <= 4001) { stuklijst = "S100104"; }
		}

		if (hoek == 37)
		{
			if (breed == 600) { tredecode = "11960"; }
			else if (breed == 650) { tredecode = "12075"; }
			else if (breed == 800) { tredecode = "10379"; }
			else if (breed == 900) { tredecode = "10380"; }
			else if (breed == 1000) { tredecode = "10381"; }
			else if (breed == 1200) { tredecode = "10382"; }
		}

		if (hoek == 42)
		{
			if (breed == 600) { tredecode = "11959"; }
			else if (breed == 650) { tredecode = "12074"; }
			else if (breed == 800) { tredecode = "10375"; }
			else if (breed == 900) { tredecode = "10376"; }
			else if (breed == 1000) { tredecode = "10377"; }
			else if (breed == 1200) { tredecode = "10378"; }
		}

		if (hoek != 00)
		{
			if (breed == 600) { supportcode = "12070"; }
			else if (breed == 650) { supportcode = ""; }
			else if (breed == 800) { supportcode = "11945"; }
			else if (breed == 900) { supportcode = "12470"; }
			else if (breed == 1000) { supportcode = "12071"; }
			else if (breed == 1200) { supportcode = "12072"; }
		}

		if (type == 0) { ssm = 1; }
		else if (type == 1) { ssm = 0; }
		else if (type == 2) { ssm = 1; }
		else if (type == 3) { ssm = 0; }
		else if (type == 4) { ssm = 2; }
		else if (type == 5) { ssm = 1; }

		decimal inputdec = Convert.ToDecimal(input);

		decimal tottrede = inputdec * treden;
		decimal totsupp = inputdec * ssm;

		if (stuklijst == "" || tredecode == "" || supportcode == "" || inputdec == 0)
		{
			return;
		}

		else
		{
			MessageBox.Show(input + " trappen" +
							"\nType " + type +
							"\n" + hoek + " graden" +
							"\n" + input1 + " mm hoog" +
							"\n" + treden + " treden" +
							"\n" + input2 + " mm breed" +
							"\n" +
							"\nArtikelcode roostertrede: " + tredecode + " - " + tottrede + " x" +
							"\nArtikelcode stairsupport plate: " + supportcode + " - " + totsupp + " x" +
							"\nStuklijstnummer: " + stuklijst + " - " + inputdec + " x"
				//			"\n" +
				//			"\nDit script heeft nog geen functie, komt er aan :)"
							, "Trebuchet");
			

			{
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", tredecode), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{

					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + tredecode);
				}
				else
				{
					ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
					rsAssemblyItem.UseDataChanges = true;
					rsAssemblyItem.AddNew();

					rsAssemblyItem.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblyItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsAssemblyItem.Fields["QUANTITY"].Value = Convert.ToDouble(tottrede);

					rsAssemblyItem.Update();

				}
			}


			{
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", supportcode), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{

					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + supportcode);
				}
				else
				{
					ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
					rsAssemblyItem.UseDataChanges = true;
					rsAssemblyItem.AddNew();

					rsAssemblyItem.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblyItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsAssemblyItem.Fields["QUANTITY"].Value = Convert.ToDouble(totsupp);

					rsAssemblyItem.Update();

				}
			}


			{
				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", stuklijst), "");
				rsSub.MoveFirst();

				if (rsSub != null && rsSub.RecordCount == 0)
				{

					MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + stuklijst);
				}
				else
				{
					ScriptRecordset rsAssemblySub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
					rsAssemblySub.UseDataChanges = true;
					rsAssemblySub.AddNew();

					rsAssemblySub.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblySub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
					rsAssemblySub.Fields["QUANTITY"].Value = Convert.ToDouble(inputdec);


					rsAssemblySub.Update();

				}
			}
		}
	}

	// M.R.v.E - 2022

}