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

		ScriptRecordset rsOrder = this.GetRecordset("R_ORDER", "", "PK_R_ORDER = " + ordernummer, "");
		rsOrder.MoveFirst();
		rsOrder.UseDataChanges = true;

		rsOrder.Fields["AKKOORDDOOR"].Value = "";
		rsOrder.Fields["DATEAKKOORD"].Value = DBNull.Value;
		rsOrder.Fields["ORDERAKKOORD"].Value = false;


		rsOrder.Update();


	}
}