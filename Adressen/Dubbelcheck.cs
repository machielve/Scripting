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


public class R_ADDRESS_Dubbel_check_Save_User : ISaveScript
{
	//When returning false fill reason to give meaningfull message to user
	public bool BeforeSave(RowData data, RowData oldData, SaveType saveType, ref string reason)
	{
		string straat = data["STREET"].ToString();
		string nummer = data["HOUSENUMBER2"].ToString();
		string toevoeging = data["ADDITIONHOUSENUMBER"].ToString();
		string postcode = data["ZIPCODE"].ToString();
		string stad = data["CITY"].ToString();
		string landcode = data["FK_COUNTRY"].ToString();
		
		
		if (!string.IsNullOrEmpty(straat) &&
			!string.IsNullOrEmpty(nummer) &&
			!string.IsNullOrEmpty(postcode) &&
			!string.IsNullOrEmpty(stad) &&
			!string.IsNullOrEmpty(landcode))
			
		{
			string query = $"STREET = '{straat}' AND HOUSENUMBER2 = '{nummer}' AND ADDITIONHOUSENUMBER = '{toevoeging}' AND ZIPCODE = '{postcode}' AND CITY = '{stad}' AND FK_COUNTRY = '{landcode}'";
			
			bool addresExists;
			try
			{
				addresExists = AddresExistsInDatabase(query);
			}
			catch (Exception dbEx)
			{
					
				reason = "Error checking for duplicate addres: " + dbEx.Message;
				return false; // Block save if database lookup fails
			}
		
			if (addresExists )
			{
					
				reason = "Error: Dit adres bestaat al in de adressen lijst (duplicate entry).";
				return false;
					
			}
		}
		

		return true;
	}






	//When returning false fill reason to give meaningfull message to user
	public bool AfterSave(RowData data, RowData oldData, SaveType saveType, ref string reason)
	{
		return true;
	}
	
	
	
	
	
	
	// Function to check if hash exists in R_DOCUMENT table
	private bool AddresExistsInDatabase(string query)
	{		
		try
		{
			Script scriptInstance = new Script(); // Or reuse the context if possible			

			ScriptRecordset rs = scriptInstance.GetRecordset("R_ADDRESS", "", query, "");

			return rs.RecordCount > 0;
		}
		catch (Exception ex)
		{
			// Optional: log error to a file, custom table, or debug output
			return false; // Treat as "not found" to avoid crashing the save
		}	
	
	}
	
}
