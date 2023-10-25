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

		//	MessageBox.Show("Synchronous code continues here.");



		MessageBox.Show("Klaar.");
	}

	public static async Task LoginAsync()
	{
		// Create an HttpClientHandler with a CookieContainer to store cookies
		var handler = new HttpClientHandler
		{
			UseCookies = true,
			CookieContainer = new CookieContainer(),
		};

		// Create an HttpClient with the handler
		var httpClient = new HttpClient(handler);

		// Set the base address
		httpClient.BaseAddress = new Uri("https://portal.deruitertransportbv.nl/Portal4uClient/");

		// Define the login URL and form data
		string loginUrl = "Login.aspx";
		var loginData = new FormUrlEncodedContent(new[]
		{
			new KeyValuePair<string, string>("tbUsername", "info@almacon.nl"),
			new KeyValuePair<string, string>("tbPassword", "2665NE"),
		});

		// Send the login POST request
		HttpResponseMessage loginResponse = await httpClient.PostAsync(loginUrl, loginData);



		if (loginResponse.IsSuccessStatusCode)
		{
			//	MessageBox.Show("Login successful!");

			// You can now use the same HttpClient to make further requests with the established session.
			string NewTransport = "Form.aspx"; //invul scherm

			var TransportData = new FormUrlEncodedContent(new[]
			{
				new KeyValuePair<string, string>("Textfield2", "Inkoopordernummer"),
				new KeyValuePair<string, string>("Textfield5", "Ophaaladres"), 
			});


			HttpResponseMessage protectedPageResponse = await httpClient.GetAsync(NewTransport);
			
			
			if (protectedPageResponse.IsSuccessStatusCode)
			{
				MessageBox.Show(protectedPageResponse.ToString());
				
				HttpResponseMessage NewTransportResponse = await httpClient.PostAsync(NewTransport, TransportData);
				
				if (NewTransportResponse.IsSuccessStatusCode)
				{
					MessageBox.Show("yeah");
				}
				
				else MessageBox.Show("bummer");
				
				
			}
			else
			{
				MessageBox.Show("Failed to access the protected page.");
			}
		}
		else
		{
			MessageBox.Show("Login failed.");
		}



	}
	
	// M.R.v.E - 2023
}