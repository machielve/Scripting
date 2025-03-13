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

		bool cb1;           // S100217, staalcon
		bool cb2;           // S100218, vloerplaten
		bool cb3;           // S100215, trappen
		bool cb4;           // S100219, leuning
		bool cb5;           // S100220, opzetplekken
		bool cb6;           // S100542, cat-ladders
		bool cb7;           // S100343, kolom bescherm
		bool cb8;           // S100569, staalcon basic

		decimal spacerQnty = 0;
		decimal shortjoistQnty = 0;




		string StuklijstId = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		ScriptRecordset rsStuklijst = this.GetRecordset("R_ASSEMBLY", "", "PK_R_ASSEMBLY= " + StuklijstId, "");
		rsStuklijst.MoveFirst();

		string tekeningnmr = rsStuklijst.Fields["DRAWINGNUMBER"].Value.ToString();
		string tekeningnmr1 = tekeningnmr.Substring(0, 5);
		string StuklijstType = rsStuklijst.Fields["CODE"].Value.ToString().Substring(0, 8);

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
					if (values[0].ToString().Substring(0, 2) == " Fa")
					{
						listA.Add("0");
					}
					else if (values[0].ToString().Substring(0, 3) == "   ")
					{
						listA.Add("0");
					}
					else listA.Add(values[0]);

					// Artcode -> naar lijst B
					if (values[1].ToString().Substring(0, 1) != "1")
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
					if (values[8].ToString().Substring(0, 5) == "     ")
					{
						listF.Add("x");
					}
					else listF.Add(values[8]);

					// Weight(regel) -> naar lijst H  
					listH.Add(values[9]);

				}
			}
		}


		// regels verwerken

		int regels = listA.Count;

		if (cb1 == true)    //Staalconstructie injectie	
		{
			StaalInput(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb2 == true)    //vloerdelen injectie	
		{
			Vloeren(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb3 == true)    //Trappen injectie	
		{
			Trappen(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb4 == true)    //Leuning injectie	
		{
			Leunings(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb5 == true)    //Opzetplekken injectie	
		{
			POPers(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb6 == true)    //Ladders injectie	
		{
			Ladders(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb7 == true)    //KolomBescherm injectie	
		{
			KolomBescherm(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}

		if (cb8 == true)    //staalconstructie basic injectie	
		{
			staalCinput(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError);
		}














		if (ListError.Count > 0 || ListSkip.Count > 0)
		{
			ErrorBuilder(ref SalesOffer, ref Filelocation, ref ErrorLocation, ref Fullpath);
			ErrorLog(ref ErrorLocation, ref ListError, ref ListSkip, ref ListLeuning, ref ErrorFile);
			MessageBox.Show(ListError.Count.ToString() + " regels in error log");

			System.Diagnostics.Process.Start(ErrorFile);
		}

		MessageBox.Show("klaar");

	}




	public void StaalInput(ref int regels, ref string StuklijstId, ref List<string> listA,
										ref List<string> listB,
										ref List<string> listC,
										ref List<string> listD,
										ref List<string> listE,
										ref List<string> listK,
										ref List<string> listL,
										ref List<string> listM,
										ref List<string> listF,
										ref List<string> listH,
										ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 3)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // Staalconstructie importeren

	public void Vloeren(ref int regels, ref string StuklijstId, ref List<string> listA,
											ref List<string> listB,
											ref List<string> listC,
											ref List<string> listD,
											ref List<string> listE,
											ref List<string> listK,
											ref List<string> listL,
											ref List<string> listM,
											ref List<string> listF,
											ref List<string> listH,
											ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 4)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // Vloerdelen importeren

	public void Trappen(ref int regels, ref string StuklijstId, ref List<string> listA,
												ref List<string> listB,
												ref List<string> listC,
												ref List<string> listD,
												ref List<string> listE,
												ref List<string> listK,
												ref List<string> listL,
												ref List<string> listM,
												ref List<string> listF,
												ref List<string> listH,
												ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 6)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // Trappen importeren

	public void Leunings(ref int regels, ref string StuklijstId, ref List<string> listA,
													ref List<string> listB,
													ref List<string> listC,
													ref List<string> listD,
													ref List<string> listE,
													ref List<string> listK,
													ref List<string> listL,
													ref List<string> listM,
													ref List<string> listF,
													ref List<string> listH,
													ref List<string> ListError)
	{		
		decimal KnieRL = 0;
		decimal KickRL = 0;
		decimal LeuningL = 0;

		for (int i = 1; i < regels; i++)
		{
			string ItemCode = listB[i].ToString();
			int phase = Convert.ToInt32((listA[i]).ToString());

			if (phase == 5 && ItemCode == "10367    ")
			{
				int aantalKr = Convert.ToInt32(listC[i].ToString());
				int LengteKr = Convert.ToInt32(listE[i].ToString());				
				decimal aantal = aantalKr * LengteKr / 1000;
				KnieRL = KnieRL + aantal;
			}

			else if (phase == 5 && ItemCode == "10370    ")
			{
				int aantalKr = Convert.ToInt32(listC[i].ToString());
				int LengteKr = Convert.ToInt32(listE[i].ToString());
				decimal aantal = aantalKr * LengteKr / 1000;
				KickRL = KickRL + aantal;
			}

			else if (phase == 5 && ItemCode == "10553    ")
			{
				int aantalLt = Convert.ToInt32(listC[i].ToString());
				int LengteLt = Convert.ToInt32(listE[i].ToString());
				decimal aantal = aantalLt * LengteLt / 1000;
				LeuningL = LeuningL + aantal;
			}

			else if (phase == 5 && (ItemCode != "10553    " || ItemCode != "10370    " || ItemCode != "10367    "))
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);
			}

		}

		// Leuning aangepast erin
		string ItemC;
		int ItemID ;
		

		if (LeuningL > 0) // handrail aantal groter als 0
		{
			decimal aantal = Math.Ceiling(LeuningL / 6);
			// profiel
			ItemC = "10553";
			ItemID = 571;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);

			//splice
			ItemC = "12258";
			ItemID = 2285;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);
		}
		
		if (KnieRL > 0) // knierail aantal groter als 0
		{
			decimal aantal = Math.Ceiling(KnieRL / 6);
			// profiel
			ItemC = "10367";
			ItemID = 385;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);

			//splice
			ItemC = "12260";
			ItemID = 2287;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);
		}
		
		if (KickRL > 0) // kickrail aantal groter als 0
		{
			decimal aantal = Math.Ceiling(KickRL / 6);
			// profiel
			ItemC = "10370";
			ItemID = 388;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);

			//splice  
			ItemC = "10371";
			ItemID = 389;

			LeuningErin(ref StuklijstId, ref aantal, ref ItemC, ref ItemID);
		}

		//	aantal meter in trefwoorden veld

		string bericht = "";
		string bericht1 = "";
		
		decimal LL1 = Math.Ceiling(LeuningL);
		string LL = Convert.ToString(LL1);
		
		decimal KN1 = Math.Ceiling(KnieRL/2);
		string KN = Convert.ToString(KN1);

		decimal KR1 = Math.Ceiling(KickRL);
		string KR = Convert.ToString(KR1);

		if (LL1 > 0) bericht1 = LL + " meter leuning."; // SHS 50 handrail

		else bericht1 = KN + " meter leuning."; // knierail als handrail
		
		
				
		string bericht3 = KR + " meter schoprand.";

		if (KR1 == 0) bericht = bericht1;

		else bericht = bericht1 + " En " + bericht3;


		ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLY", "", "PK_R_ASSEMBLY= " + StuklijstId, "");
		rsAssemblyItem.MoveFirst();
		rsAssemblyItem.UseDataChanges = true;

		rsAssemblyItem.Fields["KEYWORDS"].Value = bericht;

		rsAssemblyItem.Update();


	} // Leuning importeren

	public void POPers(ref int regels, ref string StuklijstId, ref List<string> listA,
														ref List<string> listB,
														ref List<string> listC,
														ref List<string> listD,
														ref List<string> listE,
														ref List<string> listK,
														ref List<string> listL,
														ref List<string> listM,
														ref List<string> listF,
														ref List<string> listH,
														ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 8)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // Opzetplekken importeren

	public void Ladders(ref int regels, ref string StuklijstId, ref List<string> listA,
															ref List<string> listB,
															ref List<string> listC,
															ref List<string> listD,
															ref List<string> listE,
															ref List<string> listK,
															ref List<string> listL,
															ref List<string> listM,
															ref List<string> listF,
															ref List<string> listH,
															ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 7)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // Ladders importeren

	public void KolomBescherm(ref int regels, ref string StuklijstId, ref List<string> listA,
																ref List<string> listB,
																ref List<string> listC,
																ref List<string> listD,
																ref List<string> listE,
																ref List<string> listK,
																ref List<string> listL,
																ref List<string> listM,
																ref List<string> listF,
																ref List<string> listH,
																ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 10)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);

			}
		}
	} // KolomBescherm importeren

	public void staalCinput(ref int regels, ref string StuklijstId, ref List<string> listA,
																ref List<string> listB,
																ref List<string> listC,
																ref List<string> listD,
																ref List<string> listE,
																ref List<string> listK,
																ref List<string> listL,
																ref List<string> listM,
																ref List<string> listF,
																ref List<string> listH,
																ref List<string> ListError)
	{
		for (int i = 1; i < regels; i++)
		{
			int phase = Convert.ToInt32((listA[i]).ToString());
			if (phase == 2)
			{
				knalErin(ref regels, ref StuklijstId, ref listA, ref listB, ref listC, ref listD, ref listE, ref listK, ref listL, ref listM, ref listF, ref listH, ref ListError, ref i);
			}
		}
	} // staalconstructie basic importeren













	// importeren vanaf alles

	public void knalErin(ref int regels, ref string StuklijstId, ref List<string> listA,
																	ref List<string> listB,
																	ref List<string> listC,
																	ref List<string> listD,
																	ref List<string> listE,
																	ref List<string> listK,
																	ref List<string> listL,
																	ref List<string> listM,
																	ref List<string> listF,
																	ref List<string> listH,
																	ref List<string> ListError, ref int i)
	{
		int aantal = Convert.ToInt32(listC[i]);
		decimal lengte = Convert.ToDecimal(listE[i]);
		decimal breedte = Convert.ToDecimal(listK[i]);
		decimal extraInfo = Convert.ToDecimal(listL[i]);
		string TAG = listD[i];
		string Acode = listB[i];
		string sub1 = listM[i];
		string watser = listF[i] + " - " + listD[i];

		//	MessageBox.Show(watser);

		if (Acode != "x") artinput(ref StuklijstId, ref aantal, ref Acode, ref lengte, ref breedte, ref extraInfo, ref TAG, ref watser, ref ListError);

		if (sub1 != "x") sub1input(ref StuklijstId, ref aantal, ref sub1, ref ListError);

	} //complete regel importeren	

	public void artinput(ref string StuklijstId, ref int aantal, ref String Acode, ref decimal lengte, ref decimal breedte, 
												ref decimal extraInfo, ref string TAG, ref string watser, ref List<string> ListError)
	{
		int artID;
		decimal lengte2 = 0;
		decimal breedte2 = 0;
		decimal type;

		ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", string.Format("CODE = '{0}'", Acode), "");
		rsItem.MoveFirst();

		if (rsItem.RecordCount == 0)
		{
			artID = 0;

			string ErrorRegel = "Code onbekend - Code = " + Acode  + " TAG= " + TAG + " extra - = " + watser;
			ListError.Add(ErrorRegel);

			/*
			artswap(ref Acode, ref artID, ref watser);
			*/
		}
		else
		{
			artID = Convert.ToInt32(rsItem.Fields["PK_R_ITEM"].Value.ToString());
			type = Convert.ToDecimal(rsItem.Fields["FK_ITEMUNIT"].Value.ToString());

			// Artikleeenheden Plaat en Rooster, lengte en breedte
			if (type == 10 || type == 15 || type == 30)
			{
				lengte2 = lengte / 1000;
				breedte2 = breedte / 1000;
			}

			// Artikleeenheden met een lengte maat
			else if (type == 11 || type == 17 || type == 20 || type == 23 || type == 24 || type == 31 || type == 32 || type == 36)
			{
				lengte2 = lengte / 1000;
				breedte2 = 0;
			}

			// Artikleeenheid Trapboom
			else if (type == 22 || type == 34)
			{
				lengte2 = extraInfo;
				breedte2 = 0;
			}

			// Artikleenheden welke niet hierboven gekozen worden
			else
			{
				lengte2 = 0;
				breedte2 = 0;
			}

		}

		// check max lengte/breedte

		decimal MaxL = Convert.ToDecimal(rsItem.Fields["TRADELENGTH"].Value.ToString());
		decimal MaxB = Convert.ToDecimal(rsItem.Fields["TRADEWIDTH"].Value.ToString());

		if (lengte2 > MaxL)
		{
			string ErrorRegel = "Artikel te lang - Code = " + Acode + " TAG= " + TAG + " extra - = " + watser;
			ListError.Add(ErrorRegel);		

		}

		else if (breedte2 > MaxB)
		{
			string ErrorRegel = "Artikel te breed - Code = " + Acode + " TAG= " + TAG + " extra - = " + watser;
			ListError.Add(ErrorRegel);
		}

		else
		{
			ScriptRecordset rsSlArt = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
			rsSlArt.UseDataChanges = true;
			rsSlArt.AddNew();
			rsSlArt.Fields["FK_ASSEMBLY"].Value = StuklijstId;
			rsSlArt.Fields["FK_ITEM"].Value = artID;
			rsSlArt.Fields["LENGTH"].Value = lengte2;
			rsSlArt.Fields["WIDTH"].Value = breedte2;
			rsSlArt.Fields["QUANTITY"].Value = aantal;
			rsSlArt.Fields["CAMPARAMETER"].Value = TAG;
			rsSlArt.Update();

		}
		
	} // artikel importeren

	public void sub1input(ref string StuklijstId, ref int aantal, ref String sub1, ref List<string> ListError)
	{
		string sub = sub1;

		subinput(ref StuklijstId, ref aantal, ref sub, ref ListError);


	} // sub-stuklijst 1 importeren

	public void subinput(ref string StuklijstId, ref int aantal, ref string sub, ref List<string> ListError)
	{
		int stukID;
		string stuknummer;

		ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "", string.Format("CODE = '{0}'", sub), "");
		rsSub.MoveFirst();

		if (rsSub.RecordCount == 0)
		{
			stukID = 0;
			string ErrorRegel = "Code onbekend - Code = " + sub;
			ListError.Add(ErrorRegel);
		}
		else if (rsSub.Fields["FK_WORKFLOWSTATE"].Value.ToString() != "8d7fae53-228b-4ee9-a72a-a60d0ea6c65c")
		{
			stukID = 0;
			string ErrorRegel = "Code onbruikbaar - Code = " + sub;
			ListError.Add(ErrorRegel);
		}
		else stukID = Convert.ToInt32(rsSub.Fields["PK_R_ASSEMBLY"].Value.ToString());

		if (stukID == 0)
		{
			//MessageBox.Show("Substuklijst overgeslagen");
			return;
		}

		ScriptRecordset rsSlSub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
		rsSlSub.UseDataChanges = true;
		rsSlSub.AddNew();
		rsSlSub.Fields["FK_ASSEMBLY"].Value = StuklijstId;
		rsSlSub.Fields["QUANTITY"].Value = aantal;
		rsSlSub.Fields["FK_SUBASSEMBLY"].Value = stukID;
		rsSlSub.Update();

	} // sub-stuklijsten importeren op stuklijst

	public void LeuningErin(ref string StuklijstId, ref decimal aantal, ref string ItemC, ref int ItemID)	
	{
		ScriptRecordset rsSlArt = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
		rsSlArt.UseDataChanges = true;
		rsSlArt.AddNew();
		rsSlArt.Fields["FK_ASSEMBLY"].Value = StuklijstId;
		rsSlArt.Fields["FK_ITEM"].Value = ItemID;
		rsSlArt.Fields["QUANTITY"].Value = aantal;
		rsSlArt.Update();
	
	} // gecombineerde leuning erin




	
	
	// maken van mappen en lijsten	
	
	public void MapBuilder(ref string SalesOffer, ref string Filelocation, ref string Fullpath)
	{
		string BaseFolder = @"T:\Offertes\";   

		string OfferStart = SalesOffer.Substring(0, 3);

		string OfferGroup = OfferStart + "00-" + OfferStart + @"99\";

		string rootFolder = BaseFolder + OfferGroup;

		string partialFolderName = SalesOffer;

		Fullpath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (Fullpath != null)
		{
			Filelocation = Fullpath + @"\Lijsten";

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
			ErrorLocation = Fullpath + @"\ALM_Errors";
			// Now you can use 'fullPath' to access the folder.

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



	// pop-ups voor input

	private static DialogResult ShowInputDialog1(ref string SalesOffer, ref bool cb1, 
																		ref bool cb2, 
																		ref bool cb3, 
																		ref bool cb4, 
																		ref bool cb5, 
																		ref bool cb6, 
																		ref bool cb7, 
																		ref bool cb8)
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
		cb8 = cbox8.Checked;
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