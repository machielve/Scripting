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
using System.Windows.Forms;
using System.Data;
using Ridder.Common.Script;

public class RidderScript : CommandScript
{
    /*
	
	Schaakmat import, het  programma om een BricsCAD dataextractie uit Schaakmat.xlx te importeren
	Uit te voeren vanuit een Stuklijst met de status engineering met het tabblad Almgemeen geselecteerd
    kan vanuit Schaakmat gestart worden door middel van de shortkey F7
	Geschreven door: Machiel R. van Emden mei-2022

	*/
	
	public void Execute()
	{
		string clipboardData = Clipboard.GetText();

		foreach (var myString in clipboardData.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
		{
			string[] myStrValues = myString.Split('\t');
			string slijstCode = myStrValues[0].ToString();
			string ItemCode = myStrValues[0].ToString();

			if (slijstCode.StartsWith("S1"))
			{
				ScriptRecordset rsSub = this.GetRecordset("R_ASSEMBLY", "PK_R_ASSEMBLY, DESCRIPTION, CODE", string.Format("CODE= '{0}'", slijstCode), "");
				rsSub.MoveFirst();

				if (rsSub != null && rsSub.RecordCount == 0)
				{

					MessageBox.Show("Geen overeenkomstig stuklijst kunnen vinden. Stuklijst: " + slijstCode);
				}
				else
				{
					ScriptRecordset rsAssemblySub = this.GetRecordset("R_ASSEMBLYDETAILSUBASSEMBLY", "", "PK_R_ASSEMBLYDETAILSUBASSEMBLY= -1", "");
					rsAssemblySub.UseDataChanges = true;
					rsAssemblySub.AddNew();

					rsAssemblySub.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblySub.Fields["FK_SUBASSEMBLY"].Value = rsSub.Fields["PK_R_ASSEMBLY"].Value;
					rsAssemblySub.Fields["QUANTITY"].Value = Convert.ToDouble(myStrValues[2]);

					rsAssemblySub.Update();


				}
			}
			else
			{
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "PK_R_ITEM, DESCRIPTION, CODE", string.Format("CODE = '{0}'", ItemCode), "");
				rsItem.MoveFirst();

				if (rsItem != null && rsItem.RecordCount == 0)
				{

					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + ItemCode);
				}
				else
				{
					ScriptRecordset rsAssemblyItem = this.GetRecordset("R_ASSEMBLYDETAILITEM", "", "PK_R_ASSEMBLYDETAILITEM= -1", "");
					rsAssemblyItem.UseDataChanges = true;
					rsAssemblyItem.AddNew();

					rsAssemblyItem.Fields["FK_ASSEMBLY"].Value = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue();
					rsAssemblyItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsAssemblyItem.Fields["QUANTITY"].Value = Convert.ToDouble(myStrValues[2]);
					rsAssemblyItem.Fields["LENGTH"].Value = Convert.ToDouble(myStrValues[3]);

					rsAssemblyItem.Update();
				}
			}
		}
	}

	// M.R.v.E - 2022

}

