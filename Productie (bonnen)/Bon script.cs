using System;
using ADODB;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using Ridder.Common.ADO;
using Ridder.Common.Choices;
using Ridder.Common.Search;
using Ridder.Common.Script;
using Ridder.Common.Framework;
using Ridder.Common.Login;
using System.Data;
using System.Linq;
using Ridder.Communication.Script;
using Ridder.Recordset.Extensions;


public class R_JOBORDER_Save_User : ISaveScript
{
	//When returning false fill reason to give meaningfull message to user
	public bool BeforeSave(RowData data, RowData oldData, SaveType saveType, ref string reason)
	{
		data["ADDITIONALWORK"] = false;


		string test = data["SENDDATE"].ToString();
		DateTime test2 = Convert.ToDateTime(test);
		DateTime startP = test2.AddDays(-14);


		data["MAXIMALPRODUCTIONFINISHDATE"] = data["SENDDATE"];
		data["MINIMALPRODUCTIONSTARTDATE"] = startP;
		data["PLANNEDPRODUCTIONENDDATE"] = data["SENDDATE"];
		data["PLANNEDPRODUCTIONSTARTDATE"] = startP;

		return true;
	}

	//When returning false fill reason to give meaningfull message to user
	public bool AfterSave(RowData data, RowData oldData, SaveType saveType, ref string reason)
	{
		data["ADDITIONALWORK"] = false;
		return true;
	}
}
