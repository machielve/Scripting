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
	public void Execute()
	{
		string Filelocation = "";
		string ErrorLocation = "";
		string ErrorFile = "";
		string ImportFile = "";
		string SalesOffer = "";
		string ErrorRegel = "";
		string SkipRegel = "";
		string LeuningRegel = "";
		string Fullpath = "";

		bool cb1;			// S100217, staalcon
		bool cb2;			// S100218, vloerplaten
		bool cb3;			// S100215, trappen
		bool cb4;			// S100219, leuning
		bool cb5;			// S100220, opzetplekken
		bool cb6;			// S100542, cat-ladders
		bool cb7;			// S100343, kolom bescherm
		bool cb8;			// S100569, staalcon basic

		decimal spacerQnty = 0;
		decimal shortjoistQnty = 0;




		string StuklijstId = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		ScriptRecordset rsStuklijst = this.GetRecordset("R_ASSEMBLY", "", "PK_R_ASSEMBLY= " + StuklijstId, "");
		rsStuklijst.MoveFirst();

		string tekeningnmr = rsStuklijst.Fields["DRAWINGNUMBER"].Value.ToString();
		string tekeningnmr1 = tekeningnmr.Substring(0, 5);
		string StuklijstType = rsStuklijst.Fields["CODE"].Value.ToString().substring(0,8);

		var OfferteNummer = tekeningnmr1;

		SalesOffer = OfferteNummer;



		if (StuklijstType == "S100217/")
		{
			cb1 = true;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100218/")
		{
			cb1 = false;
			cb2 = true;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100215/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = true;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100219/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = true;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100220/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = true;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100542/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = true;
			cb7 = false;
			cb8 = false;
		}
		else if (StuklijstType == "S100343/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = true;
			cb8 = false;
		}
		else if (StuklijstType == "S100569/")
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = true;
		}
		else
		{
			cb1 = false;
			cb2 = false;
			cb3 = false;
			cb4 = false;
			cb5 = false;
			cb6 = false;
			cb7 = false;
			cb8 = false;
		}





		// csv bestand ophalen

		DialogResult result = ShowInputDialog1(ref SalesOffer, ref cb1, ref cb2, ref cb3, ref cb4, ref cb5, ref cb6, ref cb7, ref cb8);

		if (result != DialogResult.OK)
		{
			MessageBox.Show("Offerte keuze afgebroken");
			return;
		}



		MapBuilder(ref SalesOffer, ref Filelocation, ref Fullpath);
		if (Filelocation == "")
		{
			return;
		}



		FileBuilder(ref Filelocation, ref ImportFile);
		if (ImportFile == "")
		{
			return;
		}

		// aanmaken van de benodigde kolommen

		List<string> listA = new List<string>();                //Phase
		List<string> listB = new List<string>();                //Artikel code
		List<string> listC = new List<string>();                //Aantal
		List<string> listD = new List<string>();                //Merk
		List<string> listE = new List<string>();                //Lengte
		List<string> listK = new List<string>();                //Breedte
		List<string> listL = new List<string>();                //Extra info
		List<string> listM = new List<string>();                //Stuklijst nummer

		List<string> listF = new List<string>();                //Profiel
		List<string> listH = new List<string>();                //Weight (regel)

		List<string> listI = new List<string>();                //
		List<string> listJ = new List<string>();                //

		List<string> ListLeuning = new List<string>();          //de leuning lijst
		List<string> ListHR = new List<string>();               //de handrail lijst
		List<string> ListKR = new List<string>();               //de knierail lijst
		List<string> ListSR = new List<string>();               //de schoprail lijst
		
		List<string> ListError = new List<string>();            //de error lijst
		List<string> ListGood = new List<string>();             //de check lijst
		List<string> ListSkip = new List<string>();             //de skip lijst






		// inlezen csv bestand en aanpassen naar de juiste kolommen

		using (StreamReader reader = new StreamReader(ImportFile))
		{
			while (!reader.EndOfStream)
			{
				var line = reader.ReadLine();
				var values = line.Split(';');

				string check1 = line.Contains(";").ToString();

				if (check1 == "True")
				{
					// Phase -> naar lijst A				
					if (values[0].ToString() == "Fase ")
					{
						listA.Add("0");
					}
					else if (values[0].ToString() == "     ")
					{
						listA.Add("0");
					}
					else listA.Add(values[0]);

					// Artcode -> naar lijst B
					if (values[1].ToString().Substring(0, 1) != "1")
					{
						listB.Add("x");
					}
					else if (values[1].ToString().Substring(0, 1) == "0")
					{
						listB.Add("x");
					}
					else listB.Add(values[1]);				

					// Aantal -> naar lijst C
					if (values[2].ToString() == "")
					{
						listC.Add("0");
					}
					else listC.Add(values[2]);

					// Merk -> naar lijst D				
					if (values[3].ToString() == "         ")
					{
						listD.Add("x");
					}
					else listD.Add(values[3]);

					// Lengte -> naar lijst E
					if (values[4].ToString() == "        ")
					{
						listE.Add("0");
					}
					else listE.Add(values[4]);

					// Breedte -> naar lijst K 
					if (values[5].ToString() == "        ")
					{
						listK.Add("0");
					}
					else listK.Add(values[5]);

					// Extra info -> naar lijst L
					if (values[6].ToString() == "           ")
					{
						listL.Add("0");
					}
					else listL.Add(values[6]);
					
					// Stuklijst nummer -> naar lijst M
					if (values[7].ToString().Substring(0, 2) != "S1")
					{
						listM.Add("x");
					}
					else listM.Add(values[7]);					

					// Profiel -> naar lijst F				  
					if (values[8].ToString().Substring(0,5) == "     ")
					{
						listF.Add("x");
					}
					else listF.Add(values[8]);

					// Weight(regel) -> naar lijst H  
					listH.Add(values[9]);

				}
			}
		}

		// regels analyseren

		int regels = listA.Count;

		for (int i = 0; i < regels; i++)
		{
			int Phase = Convert.ToInt32(listA[i].ToString());

			// fouten opsporen

			if (listB[i].ToString().Substring(0, 4) == "Art.") // header check
			{
				SkipRegel = "Header             -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListSkip.Add(SkipRegel);
			}

			else if (listB[i].ToString() == "x" && listK[i].ToString() == "x") // verschillende checks voor artikelcode = x
			{
				if (listA[i].ToString() == "x")
				{
					SkipRegel = "Geen fase          -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else if (listD[i].ToString() == "x")
				{
					SkipRegel = "Geen merk          -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else if (listK[i].ToString() == "x")
				{
					SkipRegel = "Geen stuklijst nmr -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else
				{
					ErrorRegel = "Geen Acode         -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListError.Add(ErrorRegel);
				}
			}

			// goede regels verwerken

			else if (cb8 == true)    //staalconstructie basic injectie	
			{
				staalCinput(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH);
			}





		}

		
		if (ListError.Count > 0 || ListSkip.Count > 0)
		{
			ErrorBuilder(ref SalesOffer, ref Filelocation, ref ErrorLocation, ref Fullpath);
			ErrorLog(ref ErrorLocation, ref ListError, ref ListSkip, ref ListLeuning, ref ErrorFile);
			MessageBox.Show(ListError.Count.ToString() + " regels in error log");

			System.Diagnostics.Process.Start(ErrorFile);
		}


	}

	public void staalCinput(ref int regels, ref int StuklijstId, ref List<string> listA, 
																ref List<string> listB, 
																ref List<string> listC, 
																ref List<string> listD, 
																ref List<string> listE, 
																ref List<string> listK, 
																ref List<string> listL, 
																ref List<string> listM, 
																ref List<string> listF, 
																ref List<string> listH)
	{
		for (int i = 1; i < regels; i++)
		{
			if (Convert.ToInt32(listA[i]).ToString() == "2")
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref i);

			}
		}
	}

	public void knalErin(ref int regels, ref int StuklijstId, 	ref List<string> listA,
															ref List<string> listB,
															ref List<string> listC,
															ref List<string> listD,
															ref List<string> listE,
															ref List<string> listK,
															ref List<string> listL,
															ref List<string> listM,
															ref List<string> listF,
															ref List<string> listH, ref int i)
	{
		int aantal = Convert.ToInt32(listC[i]);
		decimal lengte = Convert.ToDecimal(listE[i]);
		decimal breedte = Convert.ToDecimal(listK[i]);
		string Acode = listB[i];
		string sub1 = listM[i];
		string watser = listF[i] + " - " + listD[i];

		if (Acode != "x") artinput(ref hoofdlijstNmr, ref aantal, ref Acode, ref lengte, ref breedte, ref watser);

		if (sub1 != "x") sub1input(ref hoofdlijstNmr, ref aantal, ref sub1);

	} 

	public void artinput(ref int hoofdlijstNmr, ref int aantal, ref String Acode, ref decimal lengte, ref decimal breedte, ref string watser)
	{
		int artID;

		ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", Acode), "");
		rsItem.MoveFirst();

		if (rsItem.RecordCount == 0)
		{
			/*
			artID = 0;
			artswap(ref Acode, ref artID, ref watser);
			*/
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
	}

	public void sub1input(ref int hoofdlijstNmr, ref int aantal, ref String sub1)
	{
		string sub = sub1;

		subinput(ref hoofdlijstNmr, ref aantal, ref sub);


	} 

	public void subinput(ref int hoofdlijstNmr, ref int aantal, ref String sub)
	{
		int stukID;
		string stuknummer;

		ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "", string.Format("CODE = '{0}'", sub), "");
		rsSub.MoveFirst();

		if (rsSub.RecordCount == 0)
		{
			/*
			stuknummer = sub;
			stukID = 0;
			subswap(ref stuknummer, ref stukID);

			*/
		}
		else if (rsSub.Fields["FK_WORKFLOWSTATE"].Value.ToString() != "8d7fae53-228b-4ee9-a72a-a60d0ea6c65c")
		{
			/*
			MessageBox.Show("Stuklijst " + (rsSub.Fields["CODE"].Value.ToString()) + " status is niet beschikbaar");
			stuknummer = sub;
			stukID = 0;
			subswap(ref stuknummer, ref stukID);
			
			*/
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

	}




	public void MapBuilder(ref string SalesOffer, ref string Filelocation, ref string Fullpath)
	{
		string BaseFolder = @"T:\Offertes\";   //wacht op nieuwe offerte map structuur, Luke

		string OfferStart = SalesOffer.Substring(0, 3);

		string OfferGroup = OfferStart + "00-" + OfferStart + @"99\";

		string rootFolder = BaseFolder + OfferGroup;

		string partialFolderName = SalesOffer;

		FullPath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (FullPath != null)
		{
			Filelocation = FullPath + @"\Lijsten";

		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOffer);
			Filelocation = "";

		}
	}

	static string FindFolder(string rootFolder, string partialFolderName, string Filelocation)
	{
		try
		{
			// Get all folders in the root directory that start with the specified prefix.
			List<string> matchingFolders = Directory.GetDirectories(rootFolder)
				.Where(folder => Path.GetFileName(folder).StartsWith(partialFolderName, StringComparison.OrdinalIgnoreCase))
				.ToList();

			if (matchingFolders.Count == 1)
			{
				return matchingFolders.First(); // Return the full path of the matching folder.
			}
			
			else if (matchingFolders.Count > 1)
			{
				// Handle the case where there are multiple matching folders with the same prefix.

				DialogResult result = ShowInputDialog2(ref matchingFolders, ref Filelocation);
				
				if (result != DialogResult.OK)
				{
					MessageBox.Show("Map keuze afgebroken");
				}

				else return Filelocation;
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

		return null; // Return null if no matching folder is found.
	}

	public void FileBuilder(ref string Filelocation, ref string ImportFile)
	{
		string FileExtension = @".csv";

		ImportFile = FindFiles(Filelocation, FileExtension);

		if (ImportFile != null)
		{
			ImportFile = ImportFile;
		}
		else
		{
			MessageBox.Show("Geen bestanden gevonden op: " + Filelocation);
			ImportFile = "";

		}
	}

	static string FindFiles(string Filelocation, string FileExtension)
	{
		string ImportFile = "";
		try
		{
			// Get all files in the directory that ends with the specified suffix.
			List<string> matchingFiles = Directory.GetFiles(Filelocation)
				.Where(file => Path.GetFileName(file).EndsWith(FileExtension, StringComparison.OrdinalIgnoreCase))
				.ToList();


			if (matchingFiles.Count > 0)
			{
				// Handle the case where there are multiple matching folders with the same prefix.

				DialogResult result = ShowInputDialog3(ref matchingFiles, ref ImportFile);
				if (result != DialogResult.OK)
				{
					MessageBox.Show("Bestands keuze afgebroken");

				}
				else return ImportFile;

			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

		return null; // Return null if no matching folder is found.
	}

	public void ErrorBuilder(ref string SalesOffer, ref string Filelocation, ref string ErrorLocation, ref string Fullpath)
	{
		if (FullPath != null)
		{
			ErrorLocation = FullPath + @"\ALM_Errors";
			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOffer);

		}
	}

	public void ErrorLog(ref string ErrorLocation, ref List<String> ListError, ref List<String> ListSkip, ref List<String> ListLeuning, ref string ErrorFile)
	{
		string datum = DateTime.Now.ToString();
		string datum1 = datum.Replace(":", "_");

		ErrorFile = ErrorLocation + @"\Error - (" + datum1 + @").txt";
		try
		{
			// Write each item in the list to the file
			using (StreamWriter writer = new StreamWriter(ErrorFile))
			{
				writer.WriteLine("Error regels:");
				foreach (string item in ListError)
				{
					writer.WriteLine(item);
				}
				writer.WriteLine("");
				writer.WriteLine("Overgeslagen regels:");
				foreach (string item in ListSkip)
				{
					writer.WriteLine(item);
				}
				writer.WriteLine("");
				writer.WriteLine("Leuning regels:");
				foreach (string item in ListLeuning)
				{
					writer.WriteLine(item);
				}
				writer.WriteLine("");
				writer.WriteLine("Done");
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

	}  //creeeren van log voor overgeslagen regels





	private static DialogResult ShowInputDialog1(ref string SalesOffer, ref bool cb1, ref bool cb2,  ref bool cb3, ref bool cb4, ref bool cb5, ref bool cb6, ref bool cb7, ref bool cb8)
	{
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 2.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Tekla offerte nummer";
		inputBox.Controls.Add(label);

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(200, 25);
		textBox.Location = new System.Drawing.Point(100, 60);
		textBox.Text = SalesOffer;
		inputBox.Controls.Add(textBox);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "Accept";
		okButton.Text = "&OK";
		okButton.Size = new System.Drawing.Size(75, 25);
		okButton.Location = new System.Drawing.Point(25, 10);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "ABORT";
		cancelButton.Text = "&Cancel";
		cancelButton.Size = new System.Drawing.Size(75, 25);
		cancelButton.Location = new System.Drawing.Point(125, 10);
		inputBox.Controls.Add(cancelButton);



		System.Windows.Forms.CheckBox cbox1 = new CheckBox();
		cbox1.Location = new System.Drawing.Point(5, 100);
		cbox1.Checked = cb1;
		cbox1.Text = "Staalconstructie";
		inputBox.Controls.Add(cbox1);

		System.Windows.Forms.CheckBox cbox2 = new CheckBox();
		cbox2.Location = new System.Drawing.Point(5, 125);
		cbox2.Checked = cb2;
		cbox2.Text = "Vloerplaten";
		inputBox.Controls.Add(cbox2);

		System.Windows.Forms.CheckBox cbox3 = new CheckBox();
		cbox3.Location = new System.Drawing.Point(5, 150);
		cbox3.Checked = cb3;
		cbox3.Text = "Trappen";
		inputBox.Controls.Add(cbox3);

		System.Windows.Forms.CheckBox cbox6 = new CheckBox();
		cbox6.Location = new System.Drawing.Point(5, 175);
		cbox6.Checked = cb6;
		cbox6.Text = "Ladders";
		inputBox.Controls.Add(cbox6);

		System.Windows.Forms.CheckBox cbox4 = new CheckBox();
		cbox4.Location = new System.Drawing.Point(5, 200);
		cbox4.Checked = cb4;
		cbox4.Text = "Leuning";
		inputBox.Controls.Add(cbox4);

		System.Windows.Forms.CheckBox cbox5 = new CheckBox();
		cbox5.Location = new System.Drawing.Point(5, 225);
		cbox5.Checked = cb5;
		cbox5.Text = "Opzetplekken";
		inputBox.Controls.Add(cbox5);

		System.Windows.Forms.CheckBox cbox7 = new CheckBox();
		cbox7.Location = new System.Drawing.Point(5, 250);
		cbox7.Size = new System.Drawing.Size(200, 25);
		cbox7.Checked = cb7;
		cbox7.Text = "Kolom beschermers";
		inputBox.Controls.Add(cbox7);

		System.Windows.Forms.CheckBox cbox8 = new CheckBox();
		cbox8.Location = new System.Drawing.Point(5, 275);
		cbox8.Size = new System.Drawing.Size(200, 25);
		cbox8.Checked = cb8;
		cbox8.Text = "Staalconstructie basic";
		inputBox.Controls.Add(cbox8);


		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		SalesOffer = textBox.Text;
		cb1 = cbox1.Checked;
		cb2 = cbox2.Checked;
		cb3 = cbox3.Checked;
		cb4 = cbox4.Checked;
		cb5 = cbox5.Checked;
		cb6 = cbox6.Checked;
		cb7 = cbox7.Checked;
		return result;

	} // bevestigen of wijzigen van het offertenummer

	private static DialogResult ShowInputDialog2(ref List<string> matchingFolders, ref string Filelocation)
	{
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 2.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Tekla offerte naam";
		inputBox.Controls.Add(label);

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFolders;
		combo1.Size = new System.Drawing.Size(200, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(100, 60);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "Accept";
		okButton.Text = "&OK";
		okButton.Size = new System.Drawing.Size(75, 25);
		okButton.Location = new System.Drawing.Point(25, 10);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "ABORT";
		cancelButton.Text = "&Cancel";
		cancelButton.Size = new System.Drawing.Size(75, 25);
		cancelButton.Location = new System.Drawing.Point(125, 10);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		Filelocation = combo1.SelectedValue.ToString();
		return result;

	}  // juiste map kiezen als er meerdere mappen zijn welke beginnen met het offertenummer

	private static DialogResult ShowInputDialog3(ref List<string> matchingFiles, ref string ImportFile)
	{
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 2.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Import lijst";
		inputBox.Controls.Add(label);

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFiles;
		combo1.Size = new System.Drawing.Size(300, 25);
		combo1.DropDownWidth = 750;
		combo1.Location = new System.Drawing.Point(100, 60);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "Accept";
		okButton.Text = "&OK";
		okButton.Size = new System.Drawing.Size(75, 25);
		okButton.Location = new System.Drawing.Point(25, 10);
		inputBox.Controls.Add(okButton);

		Button cancelButton = new Button();
		cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		cancelButton.Name = "ABORT";
		cancelButton.Text = "&Cancel";
		cancelButton.Size = new System.Drawing.Size(75, 25);
		cancelButton.Location = new System.Drawing.Point(125, 10);
		inputBox.Controls.Add(cancelButton);

		inputBox.AcceptButton = okButton;
		inputBox.CancelButton = cancelButton;

		DialogResult result = inputBox.ShowDialog();

		ImportFile = combo1.SelectedValue.ToString();
		return result;

	}  // juiste bestand kiezen om te importeren vanaf de gekozen map

}