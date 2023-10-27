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
			AllowAutoRedirect = true,
		
		};

		var httpClient = new HttpClient(handler);

		// Define the login URL and form data
		string loginUrl = "https://portal.deruitertransportbv.nl/Portal4uClient/Login.aspx";


		// Send an initial GET request to the login page to capture anti-CSRF tokens
		HttpResponseMessage initialResponse = await httpClient.GetAsync(loginUrl);

		if (!initialResponse.IsSuccessStatusCode)
		{
			Console.WriteLine("Failed to load the login page. Status Code: " + initialResponse.StatusCode);
			return;
		}

		// Extract the response content as a string
		string initialResponseContent = await initialResponse.Content.ReadAsStringAsync();

		// Use regular expressions to capture anti-CSRF tokens
		string viewState = CaptureToken(initialResponseContent, "__VIEWSTATE");
		string viewStateGenerator = CaptureToken(initialResponseContent, "__VIEWSTATEGENERATOR");
		string eventValidation = CaptureToken(initialResponseContent, "__EVENTVALIDATION");
		
				
		var loginData = new FormUrlEncodedContent(new[]
		{
			new KeyValuePair<string, string>("__EVENTTARGET", ""),
			new KeyValuePair<string, string>("__EVENTARGUMENT", ""),
			new KeyValuePair<string, string>("__LASTFOCUS", ""),
			new KeyValuePair<string, string>("__VIEWSTATE", viewState),
			new KeyValuePair<string, string>("__VIEWSTATEGENERATOR", viewStateGenerator),
			new KeyValuePair<string, string>("__SCROLLPOSITIONX", "0"),
			new KeyValuePair<string, string>("__SCROLLPOSITIONY", "0"),
			new KeyValuePair<string, string>("__EVENTVALIDATION", eventValidation),
			new KeyValuePair<string, string>("tbUsername",      "***"),
			new KeyValuePair<string, string>("tbPassword",      "***"),
			new KeyValuePair<string, string>("ddlLanguage", "NL"),
			new KeyValuePair<string, string>("btnLogin", "inloggen"),
			new KeyValuePair<string, string>("hfForgotPasswordMessage", "xcc")

		});

		httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36");
		// Send the login POST request
		HttpResponseMessage loginResponse = await httpClient.PostAsync(loginUrl, loginData);


		
		/*



		// check login respons
		MessageBox.Show(loginResponse.ToString());


		
		// check all the current cookies		
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




	//	return;  // stop here during testing of login


		if (loginResponse.IsSuccessStatusCode)
		{
			string NewTransport = "https://portal.deruitertransportbv.nl/Portal4uClient/Form.aspx?PageId=1&GroupId=2&SubGroupId=6"; //invul scherm

			string inkoopnummer = "";

			string LaadDatum = "";
			string LaadNaam = "";
			string LaadAdres = "";
			string LaadPostcode = "";
			string LaadPlaats = "";
			string LaadLand = "";
			string LaadContact = "";
			string LaadTelefoon = "";

			string LosDatum = "";
			string LosNaam = "";
			string LosAdres = "";
			string LosPostcode = "";
			string LosPlaats = "";
			string LosLand = "";
			string LosContact = "";
			string LosTelefoon = "";

			string Opmerkingen = "";

			string totaalAantal = "";
			string totaalGewicht = "";


			instance.InkoopData(		ref inkoopnummer,
										ref LaadDatum, ref LaadNaam, ref LaadAdres, ref LaadPostcode, ref LaadPlaats, ref LaadLand, ref LaadContact, ref LaadTelefoon,
										ref LosDatum, ref LosNaam, ref LosAdres, ref LosPostcode, ref LosPlaats, ref LosLand, ref LosContact, ref LosTelefoon,
										ref Opmerkingen);


			instance.InkoopRegels(		ref totaalAantal, ref totaalGewicht);
			
			
			
			
			var TransportData = new FormUrlEncodedContent(new[] // create postdata
			{
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield2",      inkoopnummer),
			
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Datefield10",     LaadDatum),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield5",      LaadNaam),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield6",      LaadAdres),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield7",      LaadPostcode),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield9",      LaadPlaats),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield25",     LaadLand),

				new KeyValuePair<string, string>("ctl00$MainContentHolder$Datefield2",     	LosDatum),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield15",     LosNaam),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield16",     LosAdres),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield20",     LosPostcode),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield21",     LosPlaats),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Textfield26",     LosLand),
			
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Numberfield1",    totaalAantal),
				new KeyValuePair<string, string>("ctl00$MainContentHolder$Numberfield2", 	totaalGewicht),
			});

			HttpResponseMessage protectedPageResponse = await httpClient.GetAsync(NewTransport);  // send post data



			if (protectedPageResponse.IsSuccessStatusCode)
			{
				//MessageBox.Show(protectedPageResponse.ToString()); // check response





				HttpResponseMessage NewTransportResponse = await httpClient.PostAsync(NewTransport, TransportData);

				if (NewTransportResponse.IsSuccessStatusCode)
				{
					MessageBox.Show("Data send succesfully.");
				}

				else MessageBox.Show("Cannot send the data.");

			}
			else
			{
				MessageBox.Show("Failed to access the form page.");
			}
		}
		else
		{
			MessageBox.Show("Login failed. Status code: " + loginResponse.StatusCode);
		}



	}

	public void InkoopData(ref string inkoopnummer,
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
		LaadPlaats = "Bleiswijk";
		LaadLand = "Nederland";
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

	public void InkoopRegels(ref string totaalAantal, ref string totaalGewicht)
	{
		totaalAantal = "1";
		totaalGewicht = "1";
	
		
	
	}

	static string CaptureToken(string content, string tokenName)
	{
		string pattern = $"<input type=\"hidden\" name=\"{tokenName}\" id=\"{tokenName}\" value=\"(.*?)\" />";
		Match match = Regex.Match(content, pattern);
		if (match.Success)
		{
			return match.Groups[1].Value;
		}
		return "";
	}
	
	/*
	Als we er heen navigeren met simulated keystrokes is het:
	van homepage naar alle opdrachten = 4 x tab en dan enter
	van homepage naar nieuwe opdracht = 5 x tab en dan enter
	


*/
	

	
	
	
	
	
	
	// M.R.v.E - 2023
}