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
	public void Execute()
	{

		/*

		Default TRA regels, het  programma om de default TRA regels op een Meer magazijn order toe te voegen
		Uit te voeren vanuit een order
		Geschreven door: Machiel R. van Emden apr-2025

		*/

		var ordernummer1 = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		int ordernummer = Convert.ToInt32(ordernummer1);
		int taaknummer;
		int risiconummer;
		int maatregelnummer;
		int actiedoor = 2;          //monteur

		ScriptRecordset rsTRARegels = this.GetRecordset("U_TRAREGELS", "", "FK_ORDER = " + ordernummer1, "");
		rsTRARegels.MoveFirst();

		int regelaantal = rsTRARegels.RecordCount;

		if (regelaantal > 0)

		{
			MessageBox.Show("TRA regels al ingevuld, verwijder eerst bestaande regels.", "Error");
			return;
		}

		else
		{
			//regel 1
			taaknummer = 9;             //werken op hoogte          
			risiconummer = 39;          //valgevaar
			maatregelnummer = 115;      //gebruik harnas 
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 2 
			taaknummer = 9;             //werken op hoogte
			risiconummer = 145;         //onder elkaar werken
			maatregelnummer = 268;      //draag helm
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 3
			taaknummer = 13;            //fysiek zwaar werk
			risiconummer = 61;          //fysieke belasting tillen
			maatregelnummer = 144;      //voorlichting hulpmiddelen
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 4
			taaknummer = 13;             //fysiek zwaar werk
			risiconummer = 60;           //werkhouding
			maatregelnummer = 141;       //voorlichting tilinstructie
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 5
			taaknummer = 17;             //vallen / struikelen / beklemming
			risiconummer = 83;           //werken op hoogte
			maatregelnummer = 174;       //gebruik valharnas
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 6
			taaknummer = 18;             //vallen van voorwerpen
			risiconummer = 85;           //onder, boven elkaar werken
			maatregelnummer = 108;       //helm
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 7
			taaknummer = 18;             //vallen van voorwerpen
			risiconummer = 85;           //onder, boven elkaar werken
			maatregelnummer = 178;       //geen los gereedschap
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 8
			taaknummer = 19;             //verkeer
			risiconummer = 90;           //intern transport
			maatregelnummer = 184;       //fysieke scheiding
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 9
			taaknummer = 22;             //geluid
			risiconummer = 103;          //geluidsniveau
			maatregelnummer = 205;       //gehoorbescherming
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 10
			taaknummer = 29;             //ander werk boven/naast/onder
			risiconummer = 143;          //vallen voorwerpen
			maatregelnummer = 266;       //helm
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			//regel 11
			taaknummer = 12;             //hijsen
			risiconummer = 56;           //communicatie
			maatregelnummer = 134;       //communicatie
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

			/*
			//regel 12  NOG DOEN
			taaknummer = 33;              //beklemming
			risiconummer = 152;           //-----
			maatregelnummer = 295;        //verwijder sierraden
			RegelVuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);
			*/

			MessageBox.Show("Default regels op TRA toegevoegd.", "Default TRA regels");

		}
	}




	public void RegelVuller(ref int ordernummer, ref int taaknummer, ref int risiconummer, ref int maatregelnummer, ref int actiedoor)
	{
		ScriptRecordset rsTRAregel = this.GetRecordset("U_TRAREGELS", "", "PK_U_TRAREGELS = -1 ", "");
		rsTRAregel.MoveFirst();
		rsTRAregel.AddNew();

		rsTRAregel.Fields["FK_ORDER"].Value = ordernummer;
		rsTRAregel.Fields["FK_TRATAKEN"].Value = taaknummer;
		rsTRAregel.Fields["FK_TRARISICOS"].Value = risiconummer;
		rsTRAregel.Fields["FK_MAATREGELEN"].Value = maatregelnummer;
		rsTRAregel.Fields["TRAACTIES"].Value = actiedoor;

		rsTRAregel.Update();

	}
}