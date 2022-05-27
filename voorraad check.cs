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
using System.Data;
using Ridder.Common.Script;


public class RidderScript : JobOrderDetailItemScript
{
	public void Execute()
	{

		string bonNummer = this.Row["FK_JOBORDER"].ToString();
		string itemCode = this.Row["FK_ITEM"].ToString();

		decimal totaalNodig = 0;
		decimal aantal;

		ScriptRecordset rsArtikel = this.GetRecordset("R_ITEM", "", "PK_R_ITEM = " + itemCode, "");
		rsArtikel.MoveFirst();

		decimal voorraadIn = Convert.ToDecimal(rsArtikel.Fields["TOTALSTOCKIN"].Value.ToString());
		decimal voorraadUit = Convert.ToDecimal(rsArtikel.Fields["TOTALSTOCKOUT"].Value.ToString());
		decimal voorraadReserv = Convert.ToDecimal(rsArtikel.Fields["TOTALSTOCKRESERVATION"].Value.ToString());
		decimal voorraadInkoop = Convert.ToDecimal(rsArtikel.Fields["TOTALFUTURESTOCK"].Value.ToString());
		string artikelNummer = rsArtikel.Fields["CODE"].Value.ToString();

		decimal EcoVoorraad = (voorraadIn - voorraadUit + voorraadInkoop - voorraadReserv);

		ScriptRecordset rsBonArtikel = this.GetRecordset("R_JOBORDERDETAILITEM", "", "FK_ITEM = " + itemCode, "");
		rsBonArtikel.MoveFirst();


		while (rsBonArtikel.EOF == false)

		{
			if (rsBonArtikel.Fields["ISRESERVED"].Value.ToString() == "False" && 
				rsBonArtikel.Fields["ISRELEASED"].Value.ToString() == "False" && 
				rsBonArtikel.Fields["REGISTRATIONPATH"].Value.ToString() == "4")

			{
				 aantal = Convert.ToDecimal(rsBonArtikel.Fields["QUANTITY"].Value.ToString());
			}

			else {  aantal = 0; }

			totaalNodig += aantal;
			rsBonArtikel.MoveNext();

		}


		if (totaalNodig > EcoVoorraad)
		{
			ScriptRecordset rsToDO = this.GetRecordset("R_TODO", "", "PK_R_TODO = -1", "");
			rsToDO.MoveFirst();
			rsToDO.AddNew();

			rsToDO.Fields["DESCRIPTION"].Value = ("Artikel " + artikelNummer + " aanvullen");
			rsToDO.Fields["FK_TODOTYPE"].Value = 11;
			rsToDO.Fields["CREATOR"].Value = 13;
			rsToDO.Fields["FK_DONEBY"].Value = 13;
			rsToDO.Fields["MEMO"].Value = "Aantal op Bonnen is meer dan op voorraad";




			rsToDO.Update();


		}




	}
}