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
using System.IO;
using Microsoft.VisualBasic;


public class RidderScript : CommandScript
{
	/*
	
	Zeeslag, het Super programma om Schaakmat te vervangen
	Uit te voeren vanuit een stuklijst met de status engineering
	Geschreven door: Machiel R. van Emden mei-2022

	*/	
	
	private static DialogResult ShowInputDialog(ref string input, ref string input2, ref string input3, ref bool cb1, ref bool cb2, ref bool cb3, ref bool cb4, ref bool cb5)
	{
		System.Drawing.Size size = new System.Drawing.Size(250, 250);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.Icon = new System.Drawing.Icon(@"W:\Machiel\Ridder\Scripting\icons\ship.ico");
		inputBox.ClientSize = size;
		inputBox.Text = "Zeeslag (Schaakmat 2.0.0)";

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(size.Width - 75, 23);
		textBox.Location = new System.Drawing.Point(60, 5);
		textBox.Text = input;
		inputBox.Controls.Add(textBox);

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(size.Width - 75, 23);
		label.Location = new System.Drawing.Point(5, 5);
		label.Text = "Tekening";
		inputBox.Controls.Add(label);
	
		System.Windows.Forms.TextBox textBox2 = new TextBox();
		textBox2.Size = new System.Drawing.Size(size.Width - 75, 23);
		textBox2.Location = new System.Drawing.Point(60, 30);
		textBox2.Text = input2;
		inputBox.Controls.Add(textBox2);

		System.Windows.Forms.Label label2 = new Label();
		label2.Size = new System.Drawing.Size(size.Width - 75, 23);
		label2.Location = new System.Drawing.Point(5, 30);
		label2.Text = "Rev.";
		inputBox.Controls.Add(label2);

		System.Windows.Forms.TextBox textBox3 = new TextBox();
		textBox3.Size = new System.Drawing.Size(size.Width - 75, 23);
		textBox3.Location = new System.Drawing.Point(60, 55);
		textBox3.Text = input3;
		inputBox.Controls.Add(textBox3);

		System.Windows.Forms.Label label3 = new Label();
		label3.Size = new System.Drawing.Size(size.Width - 75, 23);
		label3.Location = new System.Drawing.Point(5, 55);
		label3.Text = "Groep  #";
		inputBox.Controls.Add(label3);


		System.Windows.Forms.CheckBox cbox1 = new CheckBox();
		cbox1.Location = new System.Drawing.Point(5, 80);
		cbox1.Checked = cb1;
		cbox1.Text = "Staalconstructie";
		inputBox.Controls.Add(cbox1);

		System.Windows.Forms.CheckBox cbox2 = new CheckBox();
		cbox2.Location = new System.Drawing.Point(5, 105);
		cbox2.Checked = cb2;
		cbox2.Text = "Vloerplaten";
		inputBox.Controls.Add(cbox2);

		System.Windows.Forms.CheckBox cbox3 = new CheckBox();
		cbox3.Location = new System.Drawing.Point(5, 130);
		cbox3.Checked = cb3;
		cbox3.Text = "Trappen";
		inputBox.Controls.Add(cbox3);

		System.Windows.Forms.CheckBox cbox4 = new CheckBox();
		cbox4.Location = new System.Drawing.Point(5, 155);
		cbox4.Checked = cb4;
		cbox4.Text = "Leuning";
		inputBox.Controls.Add(cbox4);

		System.Windows.Forms.CheckBox cbox5 = new CheckBox();
		cbox5.Location = new System.Drawing.Point(5, 180);
		cbox5.Checked = cb5;
		cbox5.Text = "Opzetplekken";
		inputBox.Controls.Add(cbox5);
		
		
		
		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height-40);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "cancelButton";
		cancelButton.Size = new System.Drawing.Size(75,23);
		cancelButton.Text = "&Cancel";
		cancelButton.Location = new System.Drawing.Point(size.Width - 80, size.Height-40);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;
		
		DialogResult result = inputBox.ShowDialog();
		input = textBox.Text;
		input2 = textBox2.Text;
		input3 = textBox3.Text;
		cb1 = cbox1.Checked;
		cb2 = cbox2.Checked;
		cb3 = cbox3.Checked;
		cb4 = cbox4.Checked;
		cb5 = cbox5.Checked;
		return result;
		
	}

	private static DialogResult ShowInputDialog2(ref string inputer, ref string wiewatwaar,ref DataTable dtTest)
	{
		System.Drawing.Size size = new System.Drawing.Size(500, 350);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Atikel reperatie";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtTest;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);
	
		System.Windows.Forms.Label label1 = new Label();
		label1.Location = new System.Drawing.Point(5, 40);
		label1.Text = wiewatwaar;
		inputBox.Controls.Add(label1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height-50);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "cancelButton";
		cancelButton.Size = new System.Drawing.Size(75, 23);
		cancelButton.Text = "&Cancel";
		cancelButton.Location = new System.Drawing.Point(size.Width - 80, size.Height-50);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();
		inputer = combo1.SelectedValue.ToString();

		return result;

	}			//Artfixer pop-up

	private static DialogResult ShowInputDialog3(ref string artswapper, ref string watser,ref DataTable dtTest)
	{
		System.Drawing.Size size = new System.Drawing.Size(500, 350);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Artikel wissel";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtTest;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
		inputBox.Controls.Add(combo1);

		System.Windows.Forms.Label label1 = new Label();
		label1.Location = new System.Drawing.Point(5, 40);
		label1.Size = new System.Drawing.Size(350, 25);
		label1.Text = watser;
		inputBox.Controls.Add(label1);

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
		artswapper = combo1.SelectedValue.ToString();

		return result;

	}			//Artswapper pop-up

	private static DialogResult ShowInputDialog4(ref string subs, ref string watdan, ref DataTable dtTest)
	{
		System.Drawing.Size size = new System.Drawing.Size(500, 350);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Substuklijst reperatie";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtTest;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		System.Windows.Forms.Label label1 = new Label();
		label1.Location = new System.Drawing.Point(5, 40);
		label1.Text = watdan;
		inputBox.Controls.Add(label1);

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
		subs = combo1.SelectedValue.ToString();

		return result;

	}               //Subfixer pop-up

	private static DialogResult ShowInputDialog5(ref string subswapper,ref DataTable dtTest)
	{
		System.Drawing.Size size = new System.Drawing.Size(500, 350);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Substuklijst wissel";

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = dtTest;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		System.Windows.Forms.Label label1 = new Label();
		label1.Location = new System.Drawing.Point(5, 40);
		label1.Text = "omschrijving";
		inputBox.Controls.Add(label1);

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
		subswapper = combo1.SelectedValue.ToString();

		return result;

	}         					//Subswapper pop-up

	public void Execute()
	{

		bool cb1;
		bool cb2;
		bool cb3;
		bool cb4;
		bool cb5;
		string groepnmr;

		string hoofdlijst = FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		int hoofdlijstNmr = Convert.ToInt32(hoofdlijst);
		
		ScriptRecordset rsHoofdlijst = this.GetRecordset("R_ASSEMBLY", "", "PK_R_ASSEMBLY = " + hoofdlijstNmr, "");
		rsHoofdlijst.MoveFirst();

		string tekeningnmr = rsHoofdlijst.Fields["DRAWINGNUMBER"].Value.ToString();
		string tekeningnmr1 = tekeningnmr.Substring(0, 5);

		if (tekeningnmr.IndexOf("#") > 1)
		{
			int loc = tekeningnmr.IndexOf("#");

			groepnmr = tekeningnmr.Substring(loc + 1);
		}
		else groepnmr = "";
		
		string hoofdlijstCode = rsHoofdlijst.Fields["CODE"].Value.ToString();

		if (hoofdlijstCode.Substring(0, 8) == "S100217/")
		{
			cb1 = true;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
		}
		else if (hoofdlijstCode.Substring(0, 8) == "S100218/")
		{
			cb1 = false;
			cb2 = true;
			cb3 = false;
			cb4 = false;
			cb5 = false;
		}
		else if (hoofdlijstCode.Substring(0, 8) == "S100215/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = true;
			cb4 = false;
			cb5 = false;
		}
		else if (hoofdlijstCode.Substring(0, 8) == "S100219/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = true;
			cb5 = false;
		}
		else if (hoofdlijstCode.Substring(0, 8) == "S100220/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = true;
		}

		else 
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
		}
		
		string input = tekeningnmr1;
		string input2 = "";
		string input3 = groepnmr;

		ShowInputDialog(ref input, ref input2, ref input3, ref cb1, ref cb2, ref cb3, ref cb4, ref cb5);
		
		string input4;
		
		if (input3 == "") 
		{
			 input4 = "";
		}
		else
		{
			 input4 = "#" + input3;
		}
	
		string tekening = input + input2 + input4;
		string bestand1 = @"W:\Almacon Ridder\Ridder Stuklijsten\DataExtracties\";
		string bestand2	= tekening;
		string bestand3 = @".csv";

		
		
		var reader = new StreamReader(File.OpenRead(bestand1+bestand2+bestand3));
		List<string> listA = new List<string>();                //Count
		List<string> listB = new List<string>();                //TAG
		List<string> listC = new List<string>();                //Name
		List<string> listD = new List<string>();                //Afmeting
		List<string> listE = new List<string>();                //Afmeting 2
		List<string> listF = new List<string>();                //Artikelcode
		List<string> listG = new List<string>();          	    //Breedte
		List<string> listH = new List<string>();                //Groep
		List<string> listI = new List<string>();                //Hoek
		List<string> listJ = new List<string>();                //Hoogte
		List<string> listK = new List<string>();                //Kwaliteit
		List<string> listL = new List<string>();          	    //Lengte
		List<string> listM = new List<string>();                //Lengte B
		List<string> listN = new List<string>();                //Optie 1
		List<string> listO = new List<string>();                //Option 1
		List<string> listP = new List<string>();                //Sterkte
		List<string> listQ = new List<string>();                //Stuklijst
		List<string> listR = new List<string>();                //Stuklijst 1
		List<string> listS = new List<string>();                //Stuklijst 2
		List<string> listT = new List<string>();                //Stuklijst 3
		List<string> listU = new List<string>();                //Type
		List<string> listV = new List<string>();                //Verdieping
		List<string> listW = new List<string>();                //Voet
		List<string> listX = new List<string>();                //W
		List<string> listY = new List<string>();                //W1
		List<string> listZ = new List<string>();                //Layer
		List<string> listAA = new List<string>();          	  	//Area poly
		List<string> listAB = new List<string>();           	//Length poly

		while (!reader.EndOfStream)
		{
		//	var header = reader.ReadLine();
			var line = reader.ReadLine();
			var values = line.Split(';');

			listA.Add(values[0]);
			listB.Add(values[1]);
				
			listC.Add(values[2]);
			listD.Add(values[3]);
			listE.Add(values[4]);
			
			if (values[5] == "") { listF.Add("-"); }
			else if (values[5].Substring(0,1) != "1" && values[5].Substring(0,1) != "2" && values[5].Substring(0,3) != "Art")
			{
				string naam = values[1];
				string maat = values[3];
				string type = values[20];
				string wiewatwaar = naam + " - " + type + " - " + maat;
				
				artfix(ref listF, ref wiewatwaar);
			}
			else { listF.Add(values[5]); }

			
			if (values[6] == "") { listG.Add("0"); }
			else if (values[6] != "Breedte")
			{
				decimal aan = Convert.ToDecimal(values[6])/1000000;
				listG.Add(Convert.ToString(aan)); 
			}
			else listG.Add(values[6]);

			
			if (values[7] == "" && values[25] == "ALM_HANDRAIL") { listH.Add("Leuning"); }
			else if (values[7] == "") { listH.Add("-"); }
			else { listH.Add(values[7]); }

			listI.Add(values[8]);
			listJ.Add(values[9]);

			if (values[10] == "") { listK.Add("-"); }
			else { listK.Add(values[10]); }

			if (values[11] == "") { listL.Add("0"); }
			else if (values[11] != "Lengte")
			{
				decimal aan1 = Convert.ToDecimal(values[11])/1000000;
				listL.Add(Convert.ToString(aan1));
			}
			else listL.Add(values[11]);
			
			listM.Add(values[12]);
			listN.Add(values[13]);
			listO.Add(values[14]);
			listP.Add(values[15]);

			if (values[16] == "") { listQ.Add("-"); }
			else if (values[16].Substring(0, 2) != "S1" && values[16].Substring(0, 3) != "Stu")
			{
				string naam = values[1];
				string maat = values[3];
				string type = values[20];
				string watdan = naam + " - " + maat + " - " + type;
				sub1fix(ref listQ, ref watdan);				
			}
			else { listQ.Add(values[16]); }
			

			if (values[17] == "") { listR.Add("-"); }
			else if (values[17].Substring(0, 2) != "S1" && values[17].Substring(0, 3) != "Stu")
			{
				string naam = values[1];
				string maat = values[3];
				string type = values[20];
				string watdan = naam + " - " + maat + " - " + type;
				sub2fix(ref listR, ref watdan);
			}
			else { listR.Add(values[17]); }
			

			if (values[18] == "") { listS.Add("-"); }
			else if (values[18].Substring(0, 2) != "S1" && values[18].Substring(0, 3) != "Stu")
			{
				string naam = values[1];
				string maat = values[3];
				string type = values[20];
				string watdan = naam + " - " + maat + " - " + type;
				sub3fix(ref listS, ref watdan);
			}
			else { listS.Add(values[18]); }

			
			if (values[19] == "") { listT.Add("-"); }
			else if (values[19].Substring(0, 2) != "S1" && values[19].Substring(0, 3) != "Stu")
			{
				string naam = values[1];
				string maat = values[3];
				string type = values[20];
				string watdan = naam + " - " + maat + " - " + type;
				sub4fix(ref listT, ref watdan);
			}
			else { listT.Add(values[19]); }

			
			listU.Add(values[20]);
			listV.Add(values[21]);
			listW.Add(values[22]);
			listX.Add(values[23]);
			listY.Add(values[24]);
			listZ.Add(values[25]);

			if (values[26] == "") { listAA.Add("0"); }
			else if (values[26] != "Area")
			{
				decimal aan2 = Convert.ToDecimal(values[26])/1000000;
				listAA.Add(Convert.ToString(aan2));
			}
			else listAA.Add(values[26]);

			if (values[27] == "") { listAB.Add("0"); }
			else if (values[27] != "Length")
			{
				decimal aan2 = Convert.ToDecimal(values[27]) / 1000000;
				listAB.Add(Convert.ToString(aan2));
			}
			else listAB.Add(values[27]);


		}

		int regels = listA.Count;

		if (cb1 == true)	//staalconstructie injectie	
		{
			staalinput(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB,ref listD,ref listU,ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT);
		}
		
		if (cb2 == true) //vloer injectie
		{
			vloerinput(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB,ref listD,ref listU,ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT);
		}
		
		if (cb3 == true) //trappen injectie
		{
			trapinput(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB,ref listD,ref listU,ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT);
		}
		
		if (cb4 == true) //leuning injectie
		{
			leuninginput(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB,ref listD,ref listU,ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT, ref listAB);
		}
		
		if (cb5 == true) //POP injectie
		{
			POPinput(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB,ref listD,ref listU,ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT);
		}

		subcombine(ref hoofdlijstNmr);
		
		MessageBox.Show("Gereed","Klaar");
		
	
	}

	public void artinput(ref int hoofdlijstNmr, ref int aantal, ref String Acode, ref decimal lengte, ref decimal breedte, ref string watser)
	{
		int artID;
		

		ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", Acode), "");
		rsItem.MoveFirst();

		if (rsItem.RecordCount == 0)
		{	
			artID = 0;
			artswap(ref Acode, ref artID, ref watser);
		}
		else artID = Convert.ToInt32(rsItem.Fields["PK_R_ITEM"].Value.ToString());

		ScriptRecordset rsSlArt = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
		rsSlArt.UseDataChanges = true;
		rsSlArt.AddNew();
		rsSlArt.Fields["FK_ASSEMBLY"].Value = hoofdlijstNmr;
		rsSlArt.Fields["FK_ITEM"].Value = artID;
		rsSlArt.Fields["LENGTH"].Value = lengte / 1000;
		rsSlArt.Fields["WIDTH"].Value = breedte / 1000;
		rsSlArt.Fields["QUANTITY"].Value = aantal;
		rsSlArt.Update();
	}		//artikel import

	public void artfix(ref List<string> listF, ref string wiewatwaar)
	{
		int go = 0;

		while (go == 0)
		{
			string inputer = "Test";
			
			ScriptRecordset rsTest = this.GetRecordset("R_ITEM", "CODE, DESCRIPTION, DRAWINGNUMBER", string.Format("UNMARKETABLE = '{0}'", false), "DESCRIPTION");
			rsTest.MoveFirst();

			DataTable dtTest = rsTest.DataTable;

			DataColumn extracolumn = new DataColumn();
			extracolumn.DataType = System.Type.GetType("System.String");
			extracolumn.ColumnName = "TOTAAL";
			extracolumn.Expression = "(CODE)+(' - ')+(DESCRIPTION)+(' - ')+(DRAWINGNUMBER)";

			dtTest.Columns.Add(extracolumn);

			ShowInputDialog2(ref inputer, ref wiewatwaar, ref dtTest);

			ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", inputer), "");
			rsItem.MoveFirst();

			if (rsItem.RecordCount > 0)
			{
				go = go + 1;
				listF.Add(inputer);
			}
			else MessageBox.Show("Artikel niet gevonden");
		}
	}						//DataEx aanvullen indien artikelnummer ontbreekt

	public void artswap(ref string Acode, ref int artID, ref string watser)
	{
		MessageBox.Show("Artikel " + Acode +" - "+watser+ " niet herkend");

		int go = 0;

		while (go == 0)
		{
			string artswapper = "Artikelcode";

			ScriptRecordset rsTest = this.GetRecordset("R_ITEM", "CODE, DESCRIPTION, DRAWINGNUMBER", string.Format("UNMARKETABLE = '{0}'", false), "DESCRIPTION");
			rsTest.MoveFirst();

			DataTable dtTest = rsTest.DataTable;

			DataColumn extracolumn = new DataColumn();
			extracolumn.DataType = System.Type.GetType("System.String");
			extracolumn.ColumnName = "TOTAAL";
			extracolumn.Expression = "(CODE)+(' - ')+(DESCRIPTION)+(' - ')+(DRAWINGNUMBER)";

			dtTest.Columns.Add(extracolumn);
			
			ShowInputDialog3(ref artswapper, ref watser, ref dtTest);

			ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", String.Format("CODE = '{0}'", artswapper), "");
			rsItem.MoveFirst();

			if (rsItem.RecordCount > 0)
			{
				go = go + 1;
				artID = Convert.ToInt32(rsItem.Fields["PK_R_ITEM"].Value.ToString());
			}
			else
			{
				artID = 0;
				MessageBox.Show("Artikel niet gevonden");
			}

		}
	}					//artikel uitwisselen indien onbekend
	
	public void sub0input(ref int hoofdlijstNmr, ref int aantalT, ref String sub0)
	{
		ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "", string.Format("CODE = '{0}'", sub0), "");
		rsSub.MoveFirst();

		ScriptRecordset rsSlSub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
		rsSlSub.UseDataChanges = true;
		rsSlSub.AddNew();
		rsSlSub.Fields["FK_ASSEMBLY"].Value = hoofdlijstNmr;
		rsSlSub.Fields["QUANTITY"].Value = aantalT;
		rsSlSub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
		rsSlSub.Update();

	}			//losse stuklijst injecteren

	public void subinput(ref int hoofdlijstNmr, ref int aantal, ref String sub)
	{
		int stukID;
		string stuknummer;

		ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "", string.Format("CODE = '{0}'", sub), "");
		rsSub.MoveFirst();

		if (rsSub.RecordCount == 0)
		{
			stuknummer = sub;
			stukID = 0;
			subswap(ref stuknummer, ref stukID);
		}
		else if (rsSub.Fields["FK_WORKFLOWSTATE"].Value.ToString() != "8d7fae53-228b-4ee9-a72a-a60d0ea6c65c")
		{
			MessageBox.Show("Stuklijst "+(rsSub.Fields["CODE"].Value.ToString())+" status is niet beschikbaar");
			stuknummer = sub;
			stukID = 0;
			subswap(ref stuknummer, ref stukID);
		}
		else stukID = Convert.ToInt32(rsSub.Fields["PK_R_ASSEMBLY"].Value.ToString());


		if (stukID == 0)
		{
			MessageBox.Show("Substuklijst overgeslagen");
			return;
		}

		ScriptRecordset rsSlSub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
		rsSlSub.UseDataChanges = true;
		rsSlSub.AddNew();
		rsSlSub.Fields["FK_ASSEMBLY"].Value = hoofdlijstNmr;
		rsSlSub.Fields["QUANTITY"].Value = aantal;
		rsSlSub.Fields["FK_SUBASSEMBLY"].Value = stukID;
		rsSlSub.Update();

	}				//stuklijsten importeren
	
	public void sub1input(ref int hoofdlijstNmr, ref int aantal, ref String sub1)
	{
		string sub = sub1;

		subinput(ref  hoofdlijstNmr, ref  aantal, ref  sub);
		

	}			//stuklijst 1 import
	
	public void sub2input(ref int hoofdlijstNmr, ref int aantal, ref String sub2)
	{
		string sub = sub2;

		subinput(ref hoofdlijstNmr, ref aantal, ref sub);

	}			//stuklijst 2 import
	
	public void sub3input(ref int hoofdlijstNmr, ref int aantal, ref String sub3)
	{
		string sub = sub3;

		subinput(ref hoofdlijstNmr, ref aantal, ref sub);
	}			//stuklijst 3 import
	
	public void sub4input(ref int hoofdlijstNmr, ref int aantal, ref String sub4)
	{
		string sub = sub4;

		subinput(ref hoofdlijstNmr, ref aantal, ref sub);

	}           	//stuklijst 4 import

	public void subfix(ref List<String> lijst, ref string watdan)
	{
		int go = 0;

		while (go == 0)
		{
			string subs = "Stuklijst nummer";
			string wfstate = "8d7fae53-228b-4ee9-a72a-a60d0ea6c65c";

			ScriptRecordset rsTest = this.GetRecordset("R_ASSEMBLY", "CODE, DESCRIPTION, KEYWORDS, REVISION", string.Format("FK_WORKFLOWSTATE = '{0}'", wfstate), "DESCRIPTION");
			rsTest.MoveFirst();

			foreach (var regel in rsTest.AsEnumerable())
			{
				if (	rsTest.Fields["DESCRIPTION"].Value.ToString() == "" 	||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"/") 	||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"\") 	||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"#")) 	rsTest.Delete();
			}
			DataTable dtTest = rsTest.DataTable;

			DataColumn extracolumn = new DataColumn();
			extracolumn.DataType = System.Type.GetType("System.String");
			extracolumn.ColumnName = "TOTAAL";
			extracolumn.Expression = "(CODE)+('- rev.')+(REVISION)+(' - ')+(DESCRIPTION)+(' - ')+(KEYWORDS)";
			dtTest.Columns.Add(extracolumn);			
			
			ShowInputDialog4(ref subs, ref watdan, ref dtTest);

			if (subs == "-")
			{
				MessageBox.Show("Substuklijst gecancelled");
				lijst.Add(subs);
				return;
			}
			ScriptRecordset rsStuklijst = this.GetRecordset("R_ASSEMBLY", "", string.Format("CODE = '{0}'", subs), "");
			rsStuklijst.MoveFirst();

			if (rsStuklijst.RecordCount > 0)
			{
				go = go + 1;
				lijst.Add(subs);
			}
			else MessageBox.Show("Stuklijst niet gevonden");
		}
	}							//Substuklijst aanvullen indien ontbreekt						
	
	public void sub1fix(ref List<string> listQ, ref string watdan)
	{
		List<string> lijst = listQ;

		subfix(ref lijst, ref watdan);
		
	}							//DataEx aanvullen indien stuklijstnummer 1 ontbreekt

	public void sub2fix(ref List<string> listR, ref string watdan)
	{
		List<string> lijst = listR;
		
		subfix(ref lijst, ref watdan);
		
	}							//DataEx aanvullen indien stuklijstnummer 2 ontbreekt

	public void sub3fix(ref List<string> listS, ref string watdan)
	{
		List<string> lijst = listS;
		
		subfix(ref lijst, ref watdan);
		
	}							//DataEx aanvullen indien stuklijstnummer 3 ontbreekt

	public void sub4fix(ref List<string> listT, ref string watdan)
	{
		List<string> lijst = listT;
		
		subfix(ref lijst, ref watdan);

	}							//DataEx aanvullen indien stuklijstnummer 4 ontbreekt
	
	public void subswap(ref string Stuknummer, ref int stukID)
	{
	//	MessageBox.Show("Stuklijst " + Stuknummer + " niet herkend");

		int go = 0;

		while (go == 0)
		{
			string subswapper = "Stuklijstnummer";
			string wfstate = "8d7fae53-228b-4ee9-a72a-a60d0ea6c65c";

			ScriptRecordset rsTest = this.GetRecordset("R_ASSEMBLY", "CODE, DESCRIPTION, KEYWORDS, REVISION", string.Format("FK_WORKFLOWSTATE = '{0}'", wfstate), "DESCRIPTION");
			rsTest.MoveFirst();

			foreach (var regel in rsTest.AsEnumerable())
			{
				if (rsTest.Fields["DESCRIPTION"].Value.ToString() == "" ||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"/") ||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"\") ||
						rsTest.Fields["CODE"].Value.ToString().Contains(@"#")) rsTest.Delete();
			}
			DataTable dtTest = rsTest.DataTable;

			DataColumn extracolumn = new DataColumn();
			extracolumn.DataType = System.Type.GetType("System.String");
			extracolumn.ColumnName = "TOTAAL";
			extracolumn.Expression = "(CODE)+('- rev.')+(REVISION)+(' - ')+(DESCRIPTION)+(' - ')+(KEYWORDS)";
			dtTest.Columns.Add(extracolumn);
						
			ShowInputDialog5(ref subswapper, ref dtTest);

			if (subswapper == "0")
			{
				go = go + 1;
				stukID = 0;
			}

			ScriptRecordset rsSublijst = this.GetRecordset("R_ASSEMBLY", "", String.Format("CODE = '{0}'", subswapper), "");
			rsSublijst.MoveFirst();

			if (rsSublijst.RecordCount > 0)
			{
				go = go + 1;
				stukID = Convert.ToInt32(rsSublijst.Fields["PK_R_ASSEMBLY"].Value.ToString());
			}

			else
			{
				stukID = 0;
				MessageBox.Show("Stuklijst niet gevonden");
			}		

		}
		
			
	}								//stuklijst uitwisselen met een andere stuklijst

	public void subcombine(ref int hoofdlijstNmr)
	{
		ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "FK_ASSEMBLY = " + hoofdlijstNmr, "FK_SUBASSEMBLY");
		rsSub.MoveFirst();
		rsSub.UseDataChanges = true;

		while (!rsSub.EOF)
		{
			int aantal = Convert.ToInt32(rsSub.Fields["QUANTITY"].Value.ToString());
			string code = rsSub.Fields["FK_SUBASSEMBLY"].Value.ToString();
			rsSub.MoveNext();

			if (rsSub.EOF)
			{
				continue;
			}

			int aantal1 = Convert.ToInt32(rsSub.Fields["QUANTITY"].Value.ToString());
			string code1 = rsSub.Fields["FK_SUBASSEMBLY"].Value.ToString();

			if (code == code1)
			{
				int totaal = aantal + aantal1;
				rsSub.Fields["QUANTITY"].Value = totaal;
				rsSub.MovePrevious();
				rsSub.Delete();
				rsSub.Update();
				rsSub.MoveFirst();
			}

			rsSub.Update();


		}

	}											//alle sub-stuklijsten combineeren als ze gelijk zijn
	
	public void staalinput(	ref int regels ,ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB, 
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT)
	{

		for (int i = 1; i < regels; i++)
		{
			if (listH[i] == "Staalconstructie")
			{
				knalErin(ref  regels, ref  hoofdlijstNmr, ref listH, ref listA, ref  listB,	ref  listD, ref  listU, ref  listL,	ref  listG,ref  listF,ref  listQ,ref  listR,ref  listS,ref  listT, ref i);
				
			}
		}
		// MessageBox.Show("staal klaar");

	}											//importeren van alle regels met groep staalconstructie

	public void vloerinput(	ref int regels, ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB,
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT)
	{
		for (int i = 1; i < regels; i++)
		{
			if (listH[i] == "Vloerplaten")
			{
				knalErin(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB, ref listD, ref listU, ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT, ref i);

			}

		}
		
		// MessageBox.Show("vloer klaar");
	}											//importeren van alle regels met groep Vloerplaten
	
	public void trapinput(	ref int regels, ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB,
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT)
	{
		for (int i = 1; i < regels; i++)
		{
			if (listH[i] == "Trappen")
			{
				knalErin(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB, ref listD, ref listU, ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT, ref i);

			}
		}

		// MessageBox.Show("trappen klaar");

	}											//importeren van alle regels met groep trappen
	
	public void leuninginput(ref int regels, ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB,
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT,
							ref List<string> listAB)
	{
		decimal leuninglengte = 0;
		
		for (int i = 1; i < regels; i++)
		{
			if (listH[i] == "Leuning")
			{
				knalErin(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB, ref listD, ref listU, ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT, ref i);

			}
		
			if (listB[i] == "Polyline" && listH[i] == "Leuning")
			{
				
				int aantalR = Convert.ToInt32(listA[i]);
				decimal lengteR = Convert.ToDecimal(listAB[i]) / 1000 / 2;
				
				decimal EXlengte = aantalR * lengteR;

				leuninglengte += EXlengte;


			}

		}
		decimal calc = leuninglengte / 6;
		decimal calc1 = Math.Ceiling(calc);
		int aantalT = Convert.ToInt32(calc1);
		String sub0 = "S100223";
		sub1input(ref hoofdlijstNmr, ref aantalT, ref sub0);

		// MessageBox.Show("leuning klaar");
	}											//importeren van alle regels met groep leuning
	
	public void POPinput(	ref int regels, ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB,
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT)
	{
		for (int i = 1; i < regels; i++)
		{
			if (listH[i] == "POP")
			{
				knalErin(ref regels, ref hoofdlijstNmr, ref listH, ref listA, ref listB, ref listD, ref listU, ref listL, ref listG, ref listF, ref listQ, ref listR, ref listS, ref listT, ref i);

			}
		}

		// MessageBox.Show("POP klaar");
	}											//importeren van alle regels met groep POP

	public void knalErin(	ref int regels, ref int hoofdlijstNmr, ref List<string> listH, ref List<string> listA, ref List<string> listB,
							ref List<string> listD,
							ref List<string> listU,
							ref List<string> listL,
							ref List<string> listG,
							ref List<string> listF,
							ref List<string> listQ,
							ref List<string> listR,
							ref List<string> listS,
							ref List<string> listT, ref int i)
	{
		int aantal = Convert.ToInt32(listA[i]);
		decimal lengte = Convert.ToDecimal(listL[i]);
		decimal breedte = Convert.ToDecimal(listG[i]);
		string Acode = listF[i];
		string sub1 = listQ[i];
		string sub2 = listR[i];
		string sub3 = listS[i];
		string sub4 = listT[i];
		string watser = listB[i] + " - " + listD[i] + " - " + listU[i];

		if (Acode != "-") artinput(ref hoofdlijstNmr, ref aantal, ref Acode, ref lengte, ref breedte, ref watser);

		if (sub1 != "-") sub1input(ref hoofdlijstNmr, ref aantal, ref sub1);

		if (sub2 != "-") sub2input(ref hoofdlijstNmr, ref aantal, ref sub2);

		if (sub3 != "-") sub3input(ref hoofdlijstNmr, ref aantal, ref sub3);

		if (sub4 != "-") sub4input(ref hoofdlijstNmr, ref aantal, ref sub4);


	}								//alle import acties aanroepen

	// M.R.v.E - 2022
	
}