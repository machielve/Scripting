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
using System.Net.Http.Headers;
using System.Text;
using System.IO;
using System.Threading.Tasks;

public class RidderScript : CommandScript
{
	/*
	de ruiter portaal, het  programma om een Ridder inkooporder door te zetten naar het online portaal van de ruiter transport.
	Uit te voeren vanuit een inkooporder met de status nieuw en de bestelwijze op webshop
	Geschreven door: Machiel R. van Emden oktober-2023
	*/
	

	
	public void Execute()
	{
		Task.Run(async () =>
		{
			await LoginAsync();
		}).Wait();

		MessageBox.Show("Klaar.");
	}

	public static async Task LoginAsync()
	{
		RidderScript instance = new RidderScript();		
		
		// Create an HttpClientHandler with a CookieContainer to store cookies
		CookieContainer cookieContainer = new CookieContainer();
		var handler = new HttpClientHandler
		{
			UseCookies = true,
			CookieContainer = cookieContainer,
		};

		// Create an HttpClient with the handler
		var httpClient = new HttpClient(handler);

		// Define the login URL and form data
		string loginUrl = "https://portal.deruitertransportbv.nl/Portal4uClient/Login.aspx";
		var loginData = new FormUrlEncodedContent(new[]
		{
			new KeyValuePair<string, string>("tbUsername", 		"info@almacon.nl"),
			new KeyValuePair<string, string>("tbPassword", 		"***"),
		});

		// Send the login POST request
		HttpResponseMessage loginResponse = await httpClient.PostAsync(loginUrl, loginData);
		
		
		
		
		/*
		
		//check all the current cookies
		MessageBox.Show(loginResponse.ToString());

		CookieCollection cookies = cookieContainer.GetCookies(new Uri("https://portal.deruitertransportbv.nl"));

		foreach (Cookie cookie in cookies)
		{
			MessageBox.Show("Cookie Name: " + cookie.Name);
			MessageBox.Show("Cookie Value: " + cookie.Value);
			MessageBox.Show("Domain: " + cookie.Domain);
			MessageBox.Show("Path: " + cookie.Path);
			MessageBox.Show("Secure: " + cookie.Secure);
			MessageBox.Show("Expires: " + cookie.Expires);
		}
		
		
		*/
		
		
		

		if (loginResponse.IsSuccessStatusCode)
		{
			// You can now use the same HttpClient to make further requests with the established session.
			string NewTransport = "https://portal.deruitertransportbv.nl/Portal4uClient/Form.aspx?PageId=1&GroupId=2&SubGroupId=6"; //invul scherm

			string inkoopnummer = "";
			
			string LaadDatum = "" ;
			string LaadNaam = "" ; 
			string LaadAdres = "" ;
			string LaadPostcode = "" ;
			string LaadPlaats = "" ;
			string LaadLand = "" ;
			string LaadContact = "" ;
			string LaadTelefoon = "" ;
			
			string LosDatum = "" ;
			string LosNaam = "";
			string LosAdres = "";
			string LosPostcode = "";
			string LosPlaats = "";
			string LosLand = "";
			string LosContact = "";
			string LosTelefoon = "";

			string Opmerkingen = "";
			

			instance.InkoopData(		ref inkoopnummer, 
										ref LaadDatum, ref LaadNaam, ref LaadAdres, ref LaadPostcode, ref LaadPlaats, ref LaadLand, ref LaadContact, ref LaadTelefoon,
										ref LosDatum, ref LosNaam, ref LosAdres, ref LosPostcode, ref LosPlaats, ref LosLand, ref LosContact, ref LosTelefoon,
										ref Opmerkingen);

		
			
		
			
			
			
			
			
			var TransportData = new FormUrlEncodedContent(new[]
			{
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield2", 		inkoopnummer),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield5", 		LaadNaam), 
			});


			HttpResponseMessage protectedPageResponse = await httpClient.GetAsync(NewTransport);
			
			
			
			if (protectedPageResponse.IsSuccessStatusCode)
			{
				MessageBox.Show(protectedPageResponse.ToString());
				
				/*
				
				HttpResponseMessage NewTransportResponse = await httpClient.PostAsync(NewTransport, TransportData);
				
				if (NewTransportResponse.IsSuccessStatusCode)
				{
					MessageBox.Show("Data send succesfully.");
				}
				
				else MessageBox.Show("Cannot send the data.");

				*/
				
				
			}
			else
			{
				MessageBox.Show("Failed to access the protected page.");
			}
		}
		else
		{
			MessageBox.Show("Login failed. Status code: " + loginResponse.StatusCode);
		}



	}

	public void InkoopData(	ref string inkoopnummer, 
							ref string LaadDatum, ref string LaadNaam, ref string LaadAdres, ref string LaadPostcode, 
							ref string LaadPlaats, ref string LaadLand, ref string LaadContact, ref string LaadTelefoon,
							ref string LosDatum, ref string LosNaam, ref string LosAdres, ref string LosPostcode,
							ref string LosPlaats, ref string LosLand, ref string LosContact, ref string LosTelefoon,
							ref string Opmerkingen)
	{
		inkoopnummer = "check";
		
		LaadDatum = "01-01-2025";
		LaadNaam = "Almacon ";
		LaadAdres = "Kristalstraat 36";
		LaadPostcode = "2665NE";
		LaadPlaats = "Bleiswijk" ;
		LaadLand = "Nederland" ;
		LaadContact = "Erik";
		LaadTelefoon = "1234";

		LosDatum = "01-02-2025";
		LosNaam = "Almacon ";
		LosAdres = "Kristalstraat 36";
		LosPostcode = "2665NE";
		LosPlaats = "Bleiswijk";
		LosLand = "Nederland";
		LosContact = "Erik";
		LosTelefoon = "1234";

		Opmerkingen = "Tralala";
		
	}
	
	
	
	
	
	
	// M.R.v.E - 2023
}