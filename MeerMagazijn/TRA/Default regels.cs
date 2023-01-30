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
	public void Execute()
	{

		var ordernummer = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();

        int taaknummer;
        int risiconummer;
        int maatregelnummer;
        int actiedoor = 3;          //monteur
        
        
        //regel 1
        taaknummer = 1;             //werken op hoogte          
        risiconummer = 1;           //valgevaar
        maatregelnummer = 1;        //gebruik harnas 
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 2
        taaknummer = 1;             //werken op hoogte
        risiconummer = 2;           //onder elkaar werken
        maatregelnummer = 2;        //draag helm
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 3
        taaknummer = 1;             //fysiek zwaar werk
        risiconummer = 2;           //fysieke belasting tillen
        maatregelnummer = 2;        //voorlichting hulpmiddelen
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 4
        taaknummer = 1;             //fysiek zwaar werk
        risiconummer = 2;           //werkhouding
        maatregelnummer = 2;        //voorlichting tilinstructie
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 5
        taaknummer = 1;             //vallen / struikelen / beklemming
        risiconummer = 2;           //werken op hoogte
        maatregelnummer = 2;        //gebruik valharnas
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 6
        taaknummer = 1;             //vallen van voorwerpen
        risiconummer = 2;           //onder, boven elkaar werken
        maatregelnummer = 2;        //helm
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 7
        taaknummer = 1;             //vallen van voorwerpen
        risiconummer = 2;           //onder, boven elkaar werken
        maatregelnummer = 2;        //geen los gereedschap
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 8
        taaknummer = 1;             //verkeer
        risiconummer = 2;           //intern transport
        maatregelnummer = 2;        //fysieke scheiding
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 9
        taaknummer = 1;             //geluid
        risiconummer = 2;           //geluidsniveau
        maatregelnummer = 2;        //gehoorbescherming
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 10
        taaknummer = 1;             //ander werk boven/naast/onder
        risiconummer = 2;           //vallen voorwerpen
        maatregelnummer = 2;        //helm
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        //regel 11
        taaknummer = 1;             //hijsen
        risiconummer = 2;           //communicatie
        maatregelnummer = 2;        //communicatie
        Regelvuller(ref ordernummer, ref taaknummer, ref risiconummer, ref maatregelnummer, ref actiedoor);

        MessageBox.show("Default regels op TRA toegevoegd.");
        

	}



    public void RegelVuller(ref int ordernummer, ref int taaknummer, ref int risiconummer, ref int maatregelnummer)
    {
        ScriptRecordset rsTRAregel = this.GetRecordset("U_TRAREGELS", "", "U_TRAREGELS = -1 ", "");
		rsTRAregel.MoveFirst();
        rsTRAregel.AddNew();

        rsTRAregel.Fields["FK_ORDER"].Value = ordernummer;
        rsTRAregel.Fields["FK_TRATAKEN"].Value = taaknummer;
        rsTRAregel.Fields["FK_TRARISICO"].Value = risiconummer;
        rsTRAregel.Fields["FK_TRAMAATREGEL"].Value = maatregelnummer;        
        rsTRAregel.Fields["TRAACTION"].Value = actiedoor;

		rsTRAregel.Update();

    }
}