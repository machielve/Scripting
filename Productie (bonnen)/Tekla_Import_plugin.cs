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
		string SalesOrder = "";
		string ErrorRegel = "";
		string SkipRegel = "";
		string LeuningRegel = "";

		decimal spacerQnty = 0;
		decimal shortjoistQnty = 0;

		string bonId = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		ScriptRecordset rsJobOrder = this.GetRecordset("R_JOBORDER", "", "PK_R_JOBORDER= " + bonId, "");
		rsJobOrder.MoveFirst();
		var OrderId = rsJobOrder.Fields["FK_ORDER"].Value.ToString();

		ScriptRecordset rsOrder = this.GetRecordset("R_ORDER", "", "PK_R_ORDER= " + OrderId, "");
		rsOrder.MoveFirst();
		SalesOrder = rsOrder.Fields["ORDERNUMBER"].Value.ToString();

		ShowInputDialog1(ref SalesOrder);
		
		

		MapBuilder(ref SalesOrder, ref Filelocation);
		FileBuilder(ref Filelocation, ref ImportFile);
		
		
		

		List<string> listA = new List<string>();                //Phase
		List<string> listB = new List<string>();                //Artikelcode
		List<string> listC = new List<string>();                //Aantal
		List<string> listD = new List<string>();                //Merk
		List<string> listE = new List<string>();                //Lengte
		List<string> listK = new List<string>();                //Breedte
		List<string> listL = new List<string>();                //Extra info
		List<string> listF = new List<string>();                //Profiel
		List<string> listH = new List<string>();                //Weight (regel)
		List<string> listI = new List<string>();                //
		List<string> listJ = new List<string>();                //

		List<string> ListError = new List<string>();            //de error lijst
		List<string> ListGood = new List<string>();             //de check lijst
		List<string> ListSkip = new List<string>();             //de skip lijst

		List<string> ListLeuning = new List<string>();          //de leuning lijst
		List<string> ListHR = new List<string>();               //de handrail lijst
		List<string> ListKR = new List<string>();               //de knierail lijst
		List<string> ListSR = new List<string>();               //de schoprail lijst


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
					if (values[0].ToString().Substring(0, 6) == "     F")
					{
						listA.Add("0");
					}
					else if (values[0].ToString() == "          ")
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
					if (values[3].ToString().Substring(0, 5) == "     ")
					{
						listD.Add("x");
					}
					else listD.Add(values[3]);

					// Lengte -> naar lijst E
					listE.Add(values[4]);

					// Breedte -> naar lijst K 
					if (values[5].ToString() == "                        ")
					{
						listK.Add("0");
					}
					else listK.Add(values[5]);

					// Extra info -> naar lijst L
					if (values[6].ToString() == " ")
					{
						listL.Add("0");
					}
					else listL.Add(values[6]);

					// Profiel -> naar lijst F				  
					listF.Add(values[7]);

					// Weight(regel) -> naar lijst H  
					listH.Add(values[8]);

				}
			}
		}


		int regels = listA.Count;

		//	MessageBox.Show(regels.ToString());

		for (int i = 0; i < regels; i++)
		{
			int Phase = Convert.ToInt32(listA[i].ToString());

			if (listB[i].ToString() == "x") // verschillende checks voor artikelcode = x
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
				else if (listD[i].ToString().Substring(0, 3) == "DUM")
				{
					SkipRegel = "Dummy             -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else
				{
					ErrorRegel = "Geen Acode         -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListError.Add(ErrorRegel);
				}
			}

			else if (listB[i].ToString().Substring(0, 4) == "Art.") // header check
			{
				SkipRegel = "Header             -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListSkip.Add(SkipRegel);
			}


			else if (Phase > 649 && Phase < 700 && listB[i].ToString().Substring(0, 5) == "10553") // && listD[i].ToString().Substring(0, 3) == "RLG" ) // Handrail regels		
			{
				LeuningRegel = "Losse regel         -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListLeuning.Add(LeuningRegel);
			}

			else if (Phase > 649 && Phase < 700 && listB[i].ToString().Substring(0, 5) == "10370" && listD[i].ToString().Substring(0, 3) == "RLG") // Kickrail regels zonder sommatie		
			{
				LeuningRegel = "Losse regel         -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListLeuning.Add(LeuningRegel);
			}

			else if (Phase > 649 && Phase < 700 && listB[i].ToString().Substring(0, 5) == "10367" && listD[i].ToString().Substring(0, 3) == "RLG") // Knierail regels zonder sommatie		
			{
				LeuningRegel = "Losse regel         -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + "         -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListLeuning.Add(LeuningRegel);
			}

			else
			{
				string ItemCode = listB[i].ToString().Substring(0, 5);
				decimal aantal = Convert.ToDecimal(listC[i].ToString());
				string fase = listA[i].ToString();
				string merk = listD[i].ToString();
				decimal lengte = Convert.ToDecimal(listE[i].ToString());
				decimal breedte = Convert.ToDecimal(listK[i].ToString());
				decimal extraInfo = Convert.ToDecimal(listL[i].ToString());
				decimal Tgewicht = Convert.ToDecimal(listH[i].ToString()) / 10;
				string extraDim = "";

				if (Phase > 649 && Phase < 700 && (ItemCode == "10553" || ItemCode == "10370" || ItemCode == "10367"))
				{
					merk = ""; // TAG veld leeg voor de sommatie regels
					aantal = Math.Ceiling(aantal / (lengte / 1000)); // aantal meters naar aantal handelslengtes					
				}

				/*
				
				if (Phase > 649 && Phase < 700 && ItemCode == "10553")
				{
					merk = ""; // TAG veld leeg voor de sommatie regels
					aantal = Math.Ceiling(aantal / (lengte / 1000)); // aantal meters naar aantal handelslengtes					
				}

				if (Phase > 649 && Phase < 700 && ItemCode == "10370")
				{
					merk = ""; // TAG veld leeg voor de sommatie regels
					aantal = Math.Ceiling(aantal / (lengte / 1000)); // aantal meters naar aantal handelslengtes					
				}

				if (Phase > 649 && Phase < 700 && ItemCode == "10367")
				{
					merk = ""; // TAG veld leeg voor de sommatie regels
					aantal = Math.Ceiling(aantal / (lengte / 1000)); // aantal meters naar aantal handelslengtes					
				}
				
				*/




				if (ItemCode == "14166") // totaal aantal joist spacers tellen
				{
					spacerQnty = spacerQnty + aantal;
				}


				// artikel info uit ridder ophalen
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", ItemCode), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0) // check op onbekend artikel
				{
					ErrorRegel = "Artikel onbekend   -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListError.Add(ErrorRegel);
				}

				else if (aantal == 0) // check op aantal = 0
				{
					ErrorRegel = "Aantal is 0        -" + "Header    -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListError.Add(ErrorRegel);
				}

				else
				{
					decimal type = Convert.ToDecimal(rsItem.Fields["FK_ITEMUNIT"].Value.ToString());
					decimal AGroup = Convert.ToDecimal(rsItem.Fields["FK_ITEMGROUP"].Value.ToString());
					int itemId = Convert.ToInt32(rsItem.Fields["PK_R_ITEM"].Value.ToString());
					int Leverwijze = 4;
					int Regtraject = Convert.ToInt32(rsItem.Fields["REGISTRATIONPATH"].Value.ToString());
					int ZaagCode = Convert.ToInt32(rsItem.Fields["DEFAULTSAWINGCODE"].Value.ToString());
					string Omschrijving = rsItem.Fields["DESCRIPTION"].Value.ToString();
					string groupId = rsItem.Fields["FK_ITEMGROUP"].Value.ToString();
					string Tekening = rsItem.Fields["DRAWINGNUMBER"].Value.ToString();
					decimal MaxL = Convert.ToDecimal(rsItem.Fields["TRADELENGTH"].Value.ToString());
					decimal MaxB = Convert.ToDecimal(rsItem.Fields["TRADEWIDTH"].Value.ToString());

					ScriptRecordset rsItemSup = this.GetRecordset("R_ITEMWAREHOUSE", "PK_R_ITEMWAREHOUSE", "MAINWAREHOUSE = 1 AND FK_ITEM = " + itemId, "");
					rsItemSup.MoveFirst();

					int magazijnId = Convert.ToInt32(rsItemSup.Fields["PK_R_ITEMWAREHOUSE"].Value.ToString());


					// Artikleeenheden Plaat en Rooster, lengte en breedte
					if (type == 10 || type == 15 || type == 30)
					{
						lengte = lengte / 1000;
						breedte = breedte / 1000;
					}

					// Artikleeenheden met een lengte maat
					else if (type == 11 || type == 17 || type == 20 || type == 23 || type == 24 || type == 31 || type == 32)
					{
						lengte = lengte / 1000;
						breedte = 0;
					}

					// Artikleeenheid Trapboom
					else if (type == 22 || type == 34)
					{
						extraDim = "L= " + lengte.ToString();
						lengte = extraInfo;
						breedte = 0;

					}

					// Artikleenheden welke niet hierboven gekozen worden
					else
					{
						lengte = 0;
						breedte = 0;
					}




					if (groupId == "119" && lengte < 1)
					{
						shortjoistQnty = shortjoistQnty + aantal;

					}



					if (breedte > MaxB) // check voor maximale Breedte
					{
						ErrorRegel = "Breedte te groot   -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
						ListError.Add(ErrorRegel);
					}

					else if (lengte > MaxL) // check voor maximale Lengte
					{
						ErrorRegel = "Lengte te groot    -" + "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
						ListError.Add(ErrorRegel);
					}

					else
					{
						// zonder Riddder update berekeningen
						if (groupId != "116" &&     // accessoires constructie
								groupId != "117" &&     // vloerdelen hout
								groupId != "119" &&     // koud gewalste liggers
								groupId != "120" &&     // bevestiging materiaal
								groupId != "125" &&     // accessoires
								groupId != "130")       // vloerdelen staal
						{
							ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
							rsJoborderItem.AddNew();

							rsJoborderItem.Fields["WEIGHT"].Value = Tgewicht;
							rsJoborderItem.Fields["FK_JOBORDER"].Value = bonId;
							rsJoborderItem.Fields["FK_ORDER"].Value = Convert.ToInt32(OrderId);
							rsJoborderItem.Fields["FK_ITEMWAREHOUSE"].Value = magazijnId;
							rsJoborderItem.Fields["DELIVERYMETHOD"].Value = Leverwijze;
							rsJoborderItem.Fields["DESCRIPTION"].Value = Omschrijving;
							rsJoborderItem.Fields["REGISTRATIONPATH"].Value = Regtraject;
							rsJoborderItem.Fields["SAWINGCODE"].Value = ZaagCode;
							rsJoborderItem.Fields["FK_ITEM"].Value = itemId;
							rsJoborderItem.Fields["QUANTITY"].Value = aantal;
							rsJoborderItem.Fields["LENGTH"].Value = Convert.ToDouble(lengte);
							rsJoborderItem.Fields["WIDTH"].Value = Convert.ToDouble(breedte);
							rsJoborderItem.Fields["CAMPARAMETER"].Value = merk;

							rsJoborderItem.Fields["TEKLA_FASE"].Value = fase;

							rsJoborderItem.Fields["DIM_W"].Value = extraDim;

							if (Tekening == "")
							{
								rsJoborderItem.Fields["CAMGEOMETRY"].Value = SalesOrder;
							}

							rsJoborderItem.Update();

						}


						// met ridder berekeningen
						else
						{
							ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
							rsJoborderItem.AddNew();

							rsJoborderItem.Fields["FK_JOBORDER"].Value = bonId;
							rsJoborderItem.Fields["FK_ORDER"].Value = Convert.ToInt32(OrderId);
							rsJoborderItem.Fields["FK_ITEMWAREHOUSE"].Value = magazijnId;
							rsJoborderItem.Fields["DELIVERYMETHOD"].Value = Leverwijze;
							rsJoborderItem.Fields["DESCRIPTION"].Value = Omschrijving;
							rsJoborderItem.Fields["REGISTRATIONPATH"].Value = Regtraject;
							rsJoborderItem.Fields["SAWINGCODE"].Value = ZaagCode;
							rsJoborderItem.Fields["LENGTH"].Value = Convert.ToDouble(lengte);
							rsJoborderItem.Fields["WIDTH"].Value = Convert.ToDouble(breedte);
							rsJoborderItem.Fields["FK_ITEM"].Value = itemId;
							rsJoborderItem.Fields["DIM_W"].Value = extraDim;
							rsJoborderItem.Fields["CAMPARAMETER"].Value = merk;

							rsJoborderItem.Fields["TEKLA_FASE"].Value = fase;

							rsJoborderItem.UseDataChanges = true;

							rsJoborderItem.Fields["QUANTITY"].Value = aantal;


							if (Tekening == "")
							{
								rsJoborderItem.Fields["CAMGEOMETRY"].Value = SalesOrder;
							}

							rsJoborderItem.Update();
						}
					}

					ListGood.Add(listD[i].ToString());

				}

			}
		}
		//	MessageBox.Show(ListGood.Count.ToString() + " regels geimporteerd");

		if (ListError.Count > 0 || ListSkip.Count > 0)
		{
			ErrorBuilder(ref SalesOrder, ref Filelocation, ref ErrorLocation);
			ErrorLog(ref ErrorLocation, ref ListError, ref ListSkip, ref ListLeuning, ref ErrorFile);
			MessageBox.Show(ListError.Count.ToString() + " regels in error log");

			System.Diagnostics.Process.Start(ErrorFile);
		}

		// hier UBW aanmaken indien nodig
		string spaceraantal = spacerQnty.ToString();
		string kortejoists = shortjoistQnty.ToString();

		if (spacerQnty > 0)
		{
			ScriptRecordset rsJoborderUBW = this.GetRecordset("R_JOBORDERDETAILOUTSOURCED", "", "PK_R_JOBORDERDETAILOUTSOURCED= -1", "");
			rsJoborderUBW.AddNew();

			rsJoborderUBW.Fields["FK_JOBORDER"].Value = Convert.ToInt32(bonId);
			rsJoborderUBW.Fields["FK_ORDER"].Value = Convert.ToInt32(OrderId);
			rsJoborderUBW.Fields["FK_OUTSOURCEDACTIVITY"].Value = 25;
			rsJoborderUBW.Fields["QUANTITY"].Value = spacerQnty;
			rsJoborderUBW.Fields["DELIVERYMETHOD"].Value = 4;

			rsJoborderUBW.Update();

			int UBWnummer = Convert.ToInt32(rsJoborderUBW.Fields["PK_R_JOBORDERDETAILOUTSOURCED"].Value.ToString());

			MessageBox.Show("U1503 toegevoegd met " + spaceraantal + " stuks. Nu koppelen aan artikelen.");

			// koppeling tabel vullen			

			// alle joist spacer ophalen van de bon
			ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "FK_JOBORDER= " + bonId + " AND FK_ITEM= 4204", "");
			rsJoborderItem.MoveFirst();

		//	MessageBox.Show(rsJoborderItem.RecordCount.ToString() + " regels te koppelen");

			while (rsJoborderItem.EOF == false)
			{
				ScriptRecordset rsJoborderKMB = this.GetRecordset("R_JOBORDERITEMOUTSOURCED", "", "PK_R_JOBORDERITEMOUTSOURCED= -1", "");
				rsJoborderKMB.AddNew();
				rsJoborderKMB.Fields["FK_JOBORDERDETAILITEM"].Value = rsJoborderItem.Fields["PK_R_JOBORDERDETAILITEM"].Value;
				rsJoborderKMB.Fields["FK_JOBORDERDETAILOUTSOURCED"].Value = UBWnummer;

				rsJoborderKMB.Update();

				rsJoborderItem.MoveNext();
			}

			MessageBox.Show("U1503 gekoppeld aan de artikelen");
		}



		if (shortjoistQnty > 0)
		{
			MessageBox.Show(kortejoists + " stuks korte koud gewalste liggers. UBW nodig voor zagen indien Sadef (U1502).");


		}





		// Einde script


	}







	public void MapBuilder(ref string SalesOrder, ref string Filelocation)
	{
		string BaseFolder = @"T:\Projecten\";

		string OrderStart = SalesOrder.Substring(0, 2);

		string OrderGroup = OrderStart + "00-" + OrderStart + @"99\";

		string rootFolder = BaseFolder + OrderGroup;

		string partialFolderName = SalesOrder; // Replace with the first 5 characters you know.

		string fullPath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (fullPath != null)
		{
			Filelocation = fullPath + @"\Lijsten";
			//	MessageBox.Show(Filelocation);

			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOrder);

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

				ShowInputDialog2(ref matchingFolders, ref Filelocation);
				return Filelocation;
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

			/*

		if (matchingFiles.Count == 1)
		{
			return matchingFiles.First(); // Return the full path of the matching folder.
		}
		else 

		*/

			if (matchingFiles.Count > 0)
			{
				// Handle the case where there are multiple matching folders with the same prefix.

				ShowInputDialog3(ref matchingFiles, ref ImportFile);

				return ImportFile;

			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

		return null; // Return null if no matching folder is found.
	}

	public void ErrorBuilder(ref string SalesOrder, ref string Filelocation, ref string ErrorLocation)
	{
		string BaseFolder = @"T:\Projecten\";

		string OrderStart = SalesOrder.Substring(0, 2);

		string OrderGroup = OrderStart + "00-" + OrderStart + @"99\";

		string rootFolder = BaseFolder + OrderGroup;

		string partialFolderName = SalesOrder; // Replace with the first 5 characters you know.

		string fullPath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (fullPath != null)
		{
			ErrorLocation = fullPath + @"\ALM_Errors";
			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOrder);

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





	private static DialogResult ShowInputDialog1(ref string SalesOrder)
	{
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 2.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Tekla project nummer";
		inputBox.Controls.Add(label);

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(200, 25);
		textBox.Location = new System.Drawing.Point(100, 60);
		textBox.Text = SalesOrder;
		inputBox.Controls.Add(textBox);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);

		inputBox.AcceptButton = okButton;

		DialogResult result = inputBox.ShowDialog();
		SalesOrder = textBox.Text;
		return result;
		
	} // bevestigen of wijzigen van het ordernummer

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
		label.Text = "Tekla project naam";
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
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);

		inputBox.AcceptButton = okButton;

		DialogResult result = inputBox.ShowDialog();
		Filelocation = combo1.SelectedValue.ToString();
		return result;
		
	}  // juiste map kiezen als er meerdere mappen zijn welke beginnen met het ordernummer

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
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);

		inputBox.AcceptButton = okButton;

		DialogResult result = inputBox.ShowDialog();
		ImportFile = combo1.SelectedValue.ToString();
		return result;
		
	}  // juiste bestand kiezen om te importeren vanaf de gekozen map

}