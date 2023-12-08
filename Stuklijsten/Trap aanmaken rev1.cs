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
	
	Trap aanmaken, het  programma om de benodigde onderdelen van een  trap in een stuklijst te stoppen
	Uit te voeren vanuit een stuklijst met de status engineering
	Geschreven door: Machiel R. van Emden jan-2023
	Update november 2023

	*/

	private static DialogResult ShowInputDialog(ref string input, ref decimal input1, ref string input2, 
												ref bool rb10, ref bool rb11, ref bool rb20, ref bool rb21,
												ref bool rb0, ref bool rb1, ref bool rb2, ref bool rb3, ref bool rb4, ref bool rb5, ref bool rb6, ref bool rb7)
	{

		System.Drawing.Size size = new System.Drawing.Size(450, 700);
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
		textBox.Size = new System.Drawing.Size(100, 25);
		textBox.Location = new System.Drawing.Point(10, 50);
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
		BoxBreed.Items.Add("650");
		BoxBreed.Items.Add("700");
		BoxBreed.Items.Add("750");
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
		groepBoxType.Size = new System.Drawing.Size(395, 250);
		groepBoxType.Location = new System.Drawing.Point(10, 225);
		groepBoxType.Text = "Trap type";


		System.Windows.Forms.RadioButton rbutton0 = new RadioButton();
		rbutton0.Size = new System.Drawing.Size(250, 25);
		rbutton0.Location = new System.Drawing.Point(10, 25);
		rbutton0.Checked = rb0;
		rbutton0.Text = "Type 0 - (BG naar platform)";
		rbutton0.Checked = true;
		groepBoxType.Controls.Add(rbutton0);

		System.Windows.Forms.RadioButton rbutton1 = new RadioButton();
		rbutton1.Size = new System.Drawing.Size(250, 25);
		rbutton1.Location = new System.Drawing.Point(10, 50);
		rbutton1.Checked = rb1;
		rbutton1.Text = "Type 1 - (BG naar landing)";
		groepBoxType.Controls.Add(rbutton1);

		System.Windows.Forms.RadioButton rbutton2 = new RadioButton();
		rbutton2.Size = new System.Drawing.Size(250, 25);
		rbutton2.Location = new System.Drawing.Point(10, 75);
		rbutton2.Checked = rb2;
		rbutton2.Text = "Type 2 - (Landing naar platform)";
		groepBoxType.Controls.Add(rbutton2);

		System.Windows.Forms.RadioButton rbutton3 = new RadioButton();
		rbutton3.Size = new System.Drawing.Size(250, 25);
		rbutton3.Location = new System.Drawing.Point(10, 100);
		rbutton3.Checked = rb3;
		rbutton3.Text = "Type 3 - (Landing naar landing)";
		groepBoxType.Controls.Add(rbutton3);

		System.Windows.Forms.RadioButton rbutton4 = new RadioButton();
		rbutton4.Size = new System.Drawing.Size(250, 25);
		rbutton4.Location = new System.Drawing.Point(10, 125);
		rbutton4.Checked = rb4;
		rbutton4.Text = "Type 4 - (Platform naar platform)";
		groepBoxType.Controls.Add(rbutton4);

		System.Windows.Forms.RadioButton rbutton5 = new RadioButton();
		rbutton5.Size = new System.Drawing.Size(250, 25);
		rbutton5.Location = new System.Drawing.Point(10, 150);
		rbutton5.Checked = rb5;
		rbutton5.Text = "Type 5 - (Platform naar landing)";
		groepBoxType.Controls.Add(rbutton5);

		System.Windows.Forms.RadioButton rbutton6 = new RadioButton();
		rbutton6.Size = new System.Drawing.Size(250, 25);
		rbutton6.Location = new System.Drawing.Point(10, 175);
		rbutton6.Checked = rb6;
		rbutton6.Text = "Type 6 - (Met stelvoet)";
		groepBoxType.Controls.Add(rbutton6);

		System.Windows.Forms.RadioButton rbutton7 = new RadioButton();
		rbutton7.Size = new System.Drawing.Size(250, 25);
		rbutton7.Location = new System.Drawing.Point(10, 200);
		rbutton7.Checked = rb6;
		rbutton7.Text = "Type 7";
		groepBoxType.Controls.Add(rbutton7);
		
		

		inputBox.Controls.Add(groepBoxType);

		//groephoek	
		GroupBox groepBoxHoek = new GroupBox();
		groepBoxHoek.Size = new System.Drawing.Size(180, 100);
		groepBoxHoek.Location = new System.Drawing.Point(10, 500);
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


		//groepschoor	
		GroupBox groepBoxSchoor = new GroupBox();
		groepBoxSchoor.Size = new System.Drawing.Size(180, 100);
		groepBoxSchoor.Location = new System.Drawing.Point(225, 500);
		groepBoxSchoor.Text = "Schoor optie";


		System.Windows.Forms.RadioButton rbutton20 = new RadioButton();
		rbutton20.Size = new System.Drawing.Size(150, 25);
		rbutton20.Location = new System.Drawing.Point(10, 25);
		rbutton20.Checked = rb20;
		rbutton20.Text = "Zonder kruisschoor";
		rbutton20.Checked = true;
		groepBoxSchoor.Controls.Add(rbutton20);

		System.Windows.Forms.RadioButton rbutton21 = new RadioButton();
		rbutton21.Size = new System.Drawing.Size(150, 25);
		rbutton21.Location = new System.Drawing.Point(10, 50);
		rbutton21.Checked = rb21;
		rbutton21.Text = "Met kruisschoor";
		groepBoxSchoor.Controls.Add(rbutton21);

		inputBox.Controls.Add(groepBoxSchoor);
		
		
		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();
		input = textBox.Text;
		input1 = textBox1.Value;
		input2 = BoxBreed.Text;

		rb10 = rbutton10.Checked;
		rb11 = rbutton11.Checked;

		rb20 = rbutton20.Checked;
		rb21 = rbutton21.Checked;

		rb0 = rbutton0.Checked;
		rb1 = rbutton1.Checked;
		rb2 = rbutton2.Checked;
		rb3 = rbutton3.Checked;
		rb4 = rbutton4.Checked;
		rb5 = rbutton5.Checked;
		rb6 = rbutton6.Checked;
		rb7 = rbutton7.Checked;

		return result;

	}

	public void Execute()
	{
		string input = "Aantal trappen";
		decimal input1 = 1;
		string input2 = "Trap breedte";
		bool rb10 = false;
		bool rb11 = false;
		
		bool rb20 = false;
		bool rb21 = false;

		bool rb0 = true;
		bool rb1 = false;
		bool rb2 = false;
		bool rb3 = false;
		bool rb4 = false;
		bool rb5 = false;
		bool rb6 = false;
		bool rb7 = false;

		ShowInputDialog(ref input, ref input1, ref input2, ref rb10, ref rb11, ref rb20, ref rb21, ref rb0, ref rb1, ref rb2, ref rb3, ref rb4, ref rb5, ref rb6, ref rb7);



		int hoek = 0;
		int type = 0;
		int ssm = 0;

		string tredecode = "";
		string supportcode = "";
		string trapcode = "";
		string trapcodeRH = "";
		string trapcodeLH = "";
		string bevessettrap = "";
		string bevessettrede = "S100337";
		string bevessetweltrede = "S100337";
		string bevessetsupplate = "S100338";
		string traptype = "";
		string Schoorlijst = "S100225";

		decimal treden = 0;
		decimal optrede;
		decimal hoog = input1;

		decimal optreden42 = Math.Round(hoog / 210, 0);
		decimal optreden37 = Math.Round(hoog / 190, 0);
		decimal breed = Int32.Parse(input2);

		if (rb10 == true)
		{
			hoek = 42;
			treden = optreden42 - 1;
			optrede = optreden42;
		}

		else if (rb11 == true)
		{
			hoek = 37;
			treden = optreden37 - 1;
			optrede = optreden37;
		}

		else
		{
			hoek = 0;
			treden = 0;
			optrede = 0;
		}

		double hoekrad = hoek * (Math.PI / 180);
		double hoogd = Convert.ToDouble(hoog);

		double lang = hoogd / (Math.Sin(hoekrad));




		//Selecteren van de juiste trapboomset
		if (rb0 == true)
		{
			type = 0;
			bevessettrap = "S100509";
			traptype = "T0";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb1 == true)
		{
			type = 1;
			bevessettrap = "S100510";
			traptype = "T1";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb2 == true)
		{
			type = 2;
			bevessettrap = "S100511";
			traptype = "T2";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb3 == true)
		{
			type = 3;
			bevessettrap = "S100512";
			traptype = "T3";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb4 == true)
		{
			type = 4;
			bevessettrap = "S100513";
			traptype = "T4";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb5 == true)
		{
			type = 5;
			bevessettrap = "S100514";
			traptype = "T5";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb6)
		{
			type = 6;
			bevessettrap = "S100522";
			traptype = "T6";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else if (rb7)
		{
			type = 7;
			bevessettrap = "S100522";
			traptype = "T7";
			if (optrede < 26)
			{
				trapcode = "10569";
				trapcodeRH = "12795";
				trapcodeLH = "12796";
			}
			else trapcode = "";
		}

		else
		{
			type = 10;
			trapcode = "";
			bevessettrap = "";
			traptype = "";
		}

		//Selecteren tredes
		if (hoek == 37)
		{
			if (breed == 600) { tredecode = "11960"; }
			else if (breed == 650) { tredecode = "12075"; }
			else if (breed == 700) { tredecode = ""; }
			else if (breed == 750) { tredecode = ""; }
			else if (breed == 800) { tredecode = "10379"; }
			else if (breed == 900) { tredecode = "10380"; }
			else if (breed == 1000) { tredecode = "10381"; }
			else if (breed == 1200) { tredecode = "10382"; }
		}

		if (hoek == 42)
		{
			if (breed == 600) { tredecode = "11959"; }
			else if (breed == 650) { tredecode = "12074"; }
			else if (breed == 700) { tredecode = ""; }
			else if (breed == 750) { tredecode = "13751"; }
			else if (breed == 800) { tredecode = "10375"; }
			else if (breed == 900) { tredecode = "10376"; }
			else if (breed == 1000) { tredecode = "10377"; }
			else if (breed == 1200) { tredecode = "10378"; }
		}

		if (hoek != 00)
		{
			if (breed == 600) { supportcode = "13750"; }
			else if (breed == 650) { supportcode = ""; }
			else if (breed == 700) { supportcode = ""; }
			else if (breed == 750) { supportcode = "13752"; }
			else if (breed == 800) { supportcode = "13453"; }
			else if (breed == 900) { supportcode = "13452"; }
			else if (breed == 1000) { supportcode = "13451"; }
			else if (breed == 1200) { supportcode = "13450"; }
		}

		if (type == 0) { ssm = 1; }
		else if (type == 1) { ssm = 0; }
		else if (type == 2) { ssm = 1; }
		else if (type == 3) { ssm = 0; }
		else if (type == 4) { ssm = 2; }
		else if (type == 5) { ssm = 1; }
		else if (type == 6) { ssm = 1; }
		else if (type == 7) { ssm = 1; }

		decimal inputdec = Convert.ToDecimal(input);

		decimal tottrede = inputdec * treden;
		decimal totsupp = inputdec * ssm;

		decimal paintarea = 1;
		int trapnummer = 1;
		int coatingnummer = 1;


		string trapcheck = trapcodeRH + " / " + trapcodeLH;
		string trapcheck1 = trapcodeRH + trapcodeLH;

		if (tredecode == "" || supportcode == "" || trapcheck1 == "" || bevessettrap == "" || inputdec == 0)
		{
			MessageBox.Show("Er ontbreken artikelcodes");
			MessageBox.Show("Stringer set = " + trapcheck +
							"\nTrap trede = " + tredecode +
							"\nWeltrede code = " + supportcode
							, "WipWapWop");
			// cancel als artikelcodes niet ingevuld zijn
			return;
		}

		else
		{
			MessageBox.Show(input + " trap(pen)" +
							"\nType " + type +
							"\n" + hoek + " graden" +
							"\n" + input1 + " mm hoog" +
							"\n" + treden + " treden" +
							"\n" + input2 + " mm breed" +
							"\n" +
							"\n" +
							"\nArtikelcode roostertrede: " + tredecode + " - " + tottrede + " x" +
							"\nArtikelcode weltrede: " + supportcode + " - " + totsupp + " x" +
							"\nArtikelcode trapboom RH: " + trapcodeRH + " - " + inputdec + " x" +
							"\nArtikelcode trapboom LH: " + trapcodeLH + " - " + inputdec + " x"
							, "Trebuchet");


			{   // treden invoegen
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

			{   //weltrede invoegen
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", supportcode), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + supportcode);
				}
				else
				{
					if (totsupp > 0)
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
			}

			{   //trapboom RH invoegen
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", trapcodeRH), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + trapcodeRH);
				}
				else
				{
					ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
					rsAssemblyItem.UseDataChanges = true;
					rsAssemblyItem.AddNew();

					rsAssemblyItem.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblyItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsAssemblyItem.Fields["LENGTH"].Value = optrede;
					rsAssemblyItem.Fields["CAMPARAMETER"].Value = traptype + " - H= " + hoog + " mm";
					rsAssemblyItem.Fields["QUANTITY"].Value = Convert.ToDouble(inputdec);

					rsAssemblyItem.Update();
				}
			}

			{   //trapboom LH invoegen
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", trapcodeLH), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + trapcodeLH);
				}
				else
				{
					ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
					rsAssemblyItem.UseDataChanges = true;
					rsAssemblyItem.AddNew();

					rsAssemblyItem.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblyItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsAssemblyItem.Fields["LENGTH"].Value = optrede;
					rsAssemblyItem.Fields["CAMPARAMETER"].Value = traptype + " - H= " + hoog + " mm";
					rsAssemblyItem.Fields["QUANTITY"].Value = Convert.ToDouble(inputdec);

					rsAssemblyItem.Update();
				}
			}

			{   // trap bevset invoegen
				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", bevessettrap), "");
				rsSub.MoveFirst();

				if (rsSub != null && rsSub.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + bevessettrap);
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

			{   // trede bevset invoegen
				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", bevessettrede), "");
				rsSub.MoveFirst();

				if (rsSub != null && rsSub.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + bevessettrede);
				}
				else
				{
					ScriptRecordset rsAssemblySub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
					rsAssemblySub.UseDataChanges = true;
					rsAssemblySub.AddNew();

					rsAssemblySub.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblySub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
					rsAssemblySub.Fields["QUANTITY"].Value = Convert.ToDouble(tottrede);

					rsAssemblySub.Update();

				}
			}

			{   // weltrede bevset invoegen
				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", bevessetweltrede), "");
				rsSub.MoveFirst();

				if (rsSub != null && rsSub.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + bevessetweltrede);
				}
				else
				{

					ScriptRecordset rsAssemblySub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
					rsAssemblySub.UseDataChanges = true;
					rsAssemblySub.AddNew();

					if (totsupp > 0)
					{
						rsAssemblySub.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
						rsAssemblySub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
						rsAssemblySub.Fields["QUANTITY"].Value = Convert.ToDouble(totsupp);

						rsAssemblySub.Update();
					}
				}
			}

			{   // Kruisschoor invoegen

				if (rb21 == true)
				{

					ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", Schoorlijst), "");
					rsSub.MoveFirst();

					if (rsSub != null && rsSub.RecordCount == 0)
					{
						MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + Schoorlijst);
					}
					else
					{
						ScriptRecordset rsAssemblySub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
						rsAssemblySub.UseDataChanges = true;
						rsAssemblySub.AddNew();

						rsAssemblySub.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
						rsAssemblySub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
						rsAssemblySub.Fields["QUANTITY"].Value = Convert.ToDouble(input);

						rsAssemblySub.Update();

					}
				}
				
			}



			MessageBox.Show("Klaar");

		}
	}

	// M.R.v.E - 2023

}