using ADODB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

using Newtonsoft.Json;

using Ridder.Common.ADO;
using Ridder.Common.Choices;
using Ridder.Common.Search;
using Ridder.Common.Script;
using Ridder.Recordset.Extensions;

public class RidderScript : CommandScript
{
	public void Execute()
	{
		// start data
		string NU = DateTime.Now.ToLongDateString();
		
		string OrderMapLocation = @"\\MM-MGMT01\Data\Almacon-data\Almacon Orders\Ridder Orders\";

		string OrderID = FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();

		ScriptRecordset rsOrder = this.GetRecordset("R_ORDER", "", "PK_R_ORDER = " + OrderID, "");
		rsOrder.MoveFirst();
		

		// relatie data ophalen		
		string RelatieID = rsOrder.Fields["FK_RELATION"].Value.ToString();

		ScriptRecordset rsRelation = this.GetRecordset("R_RELATION", "", "PK_R_RELATION = " + RelatieID, "");
		rsRelation.MoveFirst();

		string RelatieMemo = rsRelation.Fields["PLAINTEXT_WARNING"].Value.ToString();
		string RelatieNaam = rsRelation.Fields["NAME"].Value.ToString();


		// order data ophalen
		string OrderMemo = rsOrder.Fields["PLAINTEXT_INTERNALMEMO"].Value.ToString();
		string OrderNummer = rsOrder.Fields["ORDERNUMBER"].Value.ToString();
		string OrderNummerKlant = rsOrder.Fields["CUSTOMERORDERNUMBER"].Value.ToString();
		string OrderReferentie = rsOrder.Fields["REFERENCE"].Value.ToString();
		
		
		// offertes ophalen
		ScriptRecordset rsOrderregels = this.GetRecordset("R_SALESORDERDETAILASSEMBLY", "", "FK_ORDER = " + OrderID, "");
		rsOrderregels.MoveFirst();	
		
		List<string> listB = new List<string>();

		while (!rsOrderregels.EOF)
		{
			// Offerte regels
			string Offerteregel = rsOrderregels.Fields["FK_OFFERDETAILASSEMBLY"].Value.ToString();
			
			ScriptRecordset rsOfferteRegels = this.GetRecordset("R_OFFERDETAILASSEMBLY", "", "PK_R_OFFERDETAILASSEMBLY = " + Offerteregel, "");
			rsOfferteRegels.MoveFirst();

			string OfferteID = rsOfferteRegels.Fields["FK_OFFER"].Value.ToString();

			ScriptRecordset rsOfferte = this.GetRecordset("R_OFFER", "", "PK_R_OFFER = " + OfferteID, "");
			rsOfferte.MoveFirst();

			string OfferteNummer = rsOfferte.Fields["OFFERNUMBER"].Value.ToString();
			
			listB.Add(OfferteNummer);

			rsOrderregels.MoveNext();

		}
		
		listB = listB.Distinct().ToList();
		
		


		// order doc ljst maken
		ScriptRecordset rsOrderDocumenten = this.GetRecordset("R_ORDERDOCUMENT", "", "FK_ORDER = " + OrderID, "");
		rsOrderDocumenten.MoveFirst();
		
		List<string> listA = new List<string>();

		while (!rsOrderDocumenten.EOF)
		{
			// doc locaties
			string DocID = rsOrderDocumenten.Fields["FK_DOCUMENT"].Value.ToString();

			ScriptRecordset rsDocumenten = this.GetRecordset("R_DOCUMENT", "", "PK_R_DOCUMENT = " + DocID, "");
			rsDocumenten.MoveFirst();

			string DocLocatie = rsDocumenten.Fields["DOCUMENTLOCATION"].Value.ToString();

			listA.Add(DocLocatie);

			rsOrderDocumenten.MoveNext();

		}










		




		// export file maken
		var exportObject = new
		{
			Ordernummer = OrderNummer,
			KlantNaam = RelatieNaam,
			KlantOrder = OrderNummerKlant,
			KlantReferentie = OrderReferentie,
			KlantInformatie = RelatieMemo,
			OrderInformatie = OrderMemo,
			OfferteNummers = listB,
			Docs = listA

		};


		string json = JsonConvert.SerializeObject(exportObject, Newtonsoft.Json.Formatting.Indented);

		string filename = OrderNummer + " Export-" + NU;
		string extensie = @".json";

		string tempFilePath = OrderMapLocation + OrderNummer + @"\" + filename + extensie;
		System.IO.File.WriteAllText(tempFilePath, json);


		
		
		// doc toevoegen aan Ridder
		ScriptRecordset rsNewDoc = this.GetRecordset("R_DOCUMENT", "", "PK_R_DOCUMENT = -1 ", "");
		rsNewDoc.AddNew();

		rsNewDoc.Fields["DOCUMENTLOCATION"].Value = tempFilePath;
		rsNewDoc.Fields["ORIGINALLOCATION"].Value = tempFilePath;		
		rsNewDoc.Fields["DESCRIPTION"].Value = filename;
		rsNewDoc.Fields["EXTENSION"].Value = extensie;		

		rsNewDoc.Update();
		
		string NewDocID = rsNewDoc.Fields["PK_R_DOCUMENT"].Value.ToString();
		
		

		
		ScriptRecordset rsNewOrderDoc = this.GetRecordset("R_ORDERDOCUMENT", "", "PK_R_ORDERDOCUMENT = -1 ", "");
		rsNewOrderDoc.AddNew();

		rsNewOrderDoc.Fields["FK_DOCUMENT"].Value = NewDocID;
		rsNewOrderDoc.Fields["FK_ORDER"].Value = OrderID;

		rsNewOrderDoc.Update();
		
		
		MessageBox.Show("Export bestand weg geschreven naar order map");
		
		

		

	}
}