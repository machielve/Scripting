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
	
	Explodeer substuklijst, het  programma om de een substuklijstregel op te splisten in de onderliggende artikelen, d posten, en substuklijsten
	Uit te voeren vanuit een stuklijst met de status engineering
	Geschreven door: Machiel R. van Emden mei-2022

	*/
	
	public void Execute()
	{

		IRecord[] records = this.FormDataAwareFunctions.GetSelectedRecords();

		if (records.Length == 0)
			return;


		foreach (IRecord record in records)
		{
			ScriptRecordset rsSubstuklijst = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY = " + (int)record.GetPrimaryKeyValue(), "");
			rsSubstuklijst.MoveFirst();


			string stuklijstnummer = rsSubstuklijst.Fields["FK_SUBASSEMBLY"].Value.ToString();
			string stuklijstdoel = rsSubstuklijst.Fields["FK_ASSEMBLY"].Value.ToString();
			double aantal = Convert.ToDouble(rsSubstuklijst.Fields["QUANTITY"].Value.ToString());
			


			ScriptRecordset rsStuklijstItemNew = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM = -1", "");
			rsStuklijstItemNew.MoveFirst();

			ScriptRecordset rsStuklijstSubNew = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY = -1", "");
			rsStuklijstSubNew.MoveFirst();



			ScriptRecordset rsSubStuklijstItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "FK_ASSEMBLY= " + stuklijstnummer, "");
			rsSubStuklijstItem.MoveFirst();


			while (rsSubStuklijstItem.EOF == false)
			{


				string itemCode = rsSubStuklijstItem.Fields["FK_ITEM"].Value.ToString();
				double itemAantal = Convert.ToDouble(rsSubStuklijstItem.Fields["QUANTITY"].Value.ToString());

				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "", "PK_R_ITEM= " + itemCode, "");
				rsItem.MoveFirst();

				double totaalItem = itemAantal * aantal;
				double lengte = Convert.ToDouble(rsSubStuklijstItem.Fields["LENGTH"].Value.ToString());

				rsStuklijstItemNew.AddNew();
				rsStuklijstItemNew.Fields["FK_ASSEMBLY"].Value = stuklijstdoel;
				rsStuklijstItemNew.Fields["FK_ITEM"].Value = itemCode;
				rsStuklijstItemNew.Fields["QUANTITY"].Value = totaalItem;
				rsStuklijstItemNew.Fields["LENGTH"].Value = lengte;
				rsStuklijstItemNew.Fields["DESCRIPTION"].Value = rsItem.Fields["DESCRIPTION"].Value.ToString();
				rsStuklijstItemNew.Update();

				rsSubStuklijstItem.MoveNext();

			}



			ScriptRecordset rsSubStuklijstSub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "FK_ASSEMBLY= " + stuklijstnummer, "");
			rsSubStuklijstSub.MoveFirst();


			while (rsSubStuklijstSub.EOF == false)
			{
				string SubCode = rsSubStuklijstSub.Fields["FK_SUBASSEMBLY"].Value.ToString();
				double SubAantal = Convert.ToDouble(rsSubStuklijstSub.Fields["QUANTITY"].Value.ToString());

				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "", "PK_R_ASSEMBLY= " + SubCode, "");
				rsSub.MoveFirst();

				double totaalItem = SubAantal * aantal;

				rsStuklijstSubNew.AddNew();
				rsStuklijstSubNew.Fields["FK_ASSEMBLY"].Value = stuklijstdoel;
				rsStuklijstSubNew.Fields["FK_SUBASSEMBLY"].Value = SubCode;
				rsStuklijstSubNew.Fields["QUANTITY"].Value = totaalItem;
				rsStuklijstSubNew.Fields["DESCRIPTION"].Value = rsSub.Fields["DESCRIPTION"].Value.ToString();
				rsStuklijstSubNew.Update();

				rsSubStuklijstSub.MoveNext();

			}


			rsSubstuklijst.Delete();

		}

		this.FormDataAwareFunctions.Refres();


	}

	// M.R.v.E - 2022
	
}