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
using System.Net;
using System.Net.Http;

public class RidderScript : CommandScript
{
	/*
	
	De ruiter portaal, het  programma om een ridder inkoop order door te schuiven naar een inkoop via het de ruiter portaal
	Uit te voeren vanuit een inkooporder met de status nieuw
	Geschreven door: Machiel R. van Emden oktober-2023

	*/
	
	public void Execute()
	{

        string website = "https://portal.deruitertransportbv.nl/Portal4uClient/Login.aspx";

        string user = "info@almacon.nl";

        string password = "2665NN"; //nog aanpassen


        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(website);
        CookieContainer cookies = new CookieContainer();







    }


	// M.R.v.E - 2023

}
