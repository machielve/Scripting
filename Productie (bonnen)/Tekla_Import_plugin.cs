using ADODB;
using System;
using System.IO;
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
		string Filelocation = "";
		string ErrorLocation = "";
		string ImportFile = "";
		string SalesOrder = "";
		string ErrorRegel = "" ;
		string SkipRegel = "" ;

		
		string bonId = this.FormDataAwareFunctions.CurrentRecord.GetPrimaryKeyValue().ToString();
		ScriptRecordset rsJobOrder = this.GetRecordset("R_JOBORDER", "", "PK_R_JOBORDER= " + bonId, "");
		rsJobOrder.MoveFirst();
		string OrderId = rsJobOrder.Fields["FK_ORDER"].ToString();		
		ScriptRecordset rsOrder = this.GetRecordset("R_ORDER", "", "PK_R_ORDER= " + OrderId, "");
		rsOrder.MoveFirst();
		SalesOrder = rsOrder.Fields["ORDERNUMBER"].ToString();
				
		
		SalesOrder = "6830";
		
		
		ShowInputDialog1(ref SalesOrder);

		MapBuilder(ref SalesOrder, ref Filelocation);
		FileBuilder(ref Filelocation, ref ImportFile);
	
		var reader = new StreamReader(File.OpenRead(ImportFile));
		List<string> listA = new List<string>();                //Phase
		List<string> listB = new List<string>();                //Artikelcode
		List<string> listC = new List<string>();                //Aantal
		List<string> listD = new List<string>();                //Merk
		List<string> listE = new List<string>();                //Lengte
		List<string> listF = new List<string>();                //Profiel
		List<string> listG = new List<string>();                //Weight (stuk)
		List<string> listH = new List<string>();                //Weight (regel)
		List<string> listI = new List<string>();                //
		List<string> listJ = new List<string>();                //

		List<string> ListError = new List<string>();            //de error lijst
		List<string> ListGood = new List<string>();             //de check lijst
		List<string> ListSkip = new List<string>();             //de skip lijst



		while (!reader.EndOfStream)
		{
			var line = reader.ReadLine();
			var values = line.Split(';');

			string check1 = line.Contains(";").ToString();

			if (check1 == "True")
			{
				//MessageBox.Show(line.ToString());
				
				
				//kolom A phase -> naar lijst A				
				if (values[0].ToString().Substring(0, 6) == "     F")
				{
					listA.Add("x");
				}
				else listA.Add(values[0]);
				
				//kolom B artcode -> naar lijst B
				if (values[1].ToString().Substring(0,1) != "1")
				{
					listB.Add("x");
				}
				else if (values[1].ToString().Substring(0,1) == "0")
				{
					listB.Add("x");
				}
				else listB.Add(values[1]);

				//kolom C aantal -> naar lijst C
				if (values[2].ToString() == "")
				{
					listC.Add("0");
				}
				else listC.Add(values[2]);
				
   				//kolom D merk -> naar lijst D				
				if (values[3].ToString().Substring(0, 5) == "     ")
				{
					listD.Add("x");
				}
				else listD.Add(values[3]);  
				
				//kolom E lengte -> naar lijst E
				listE.Add(values[4]); 
				
				//kolom F profiel -> naar lijst F				  
				listF.Add(values[5]);   
				
				//kolom G weight (stuk) -> naar lijst G
				listG.Add(values[6]); 
				
				//kolom H weight(regel) -> naar lijst H  
				listH.Add(values[7]);   


			}
			
		}
		
		int regels = listA.Count;


		
		

		for (int i = 0; i < regels; i++)
		{
			if (listB[i].ToString() == "x")
			{
				if (listA[i].ToString() == "x")
				{
					SkipRegel = "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else if (listD[i].ToString() == "x")
				{
					SkipRegel = "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListSkip.Add(SkipRegel);
				}
				else
				{
					ErrorRegel = "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
					ListError.Add(ErrorRegel);
				}		
			}
			
			else if (listB[i].ToString().Substring(0,4) == "Art.")
			{
				SkipRegel = "Fase= " + listA[i].ToString() + "Art.code= " + listB[i].ToString() + " -Merk= " + listD[i].ToString() + " -Profiel= " + listF[i].ToString();
				ListSkip.Add(SkipRegel);
			}
			
			
			else 
			{
				string ItemCode = listB[i].ToString();
				decimal aantal = Convert.ToDecimal(listC[i].ToString());
				string fase = listA[i].ToString();
				string merk = listD[i].ToString();
				decimal lengte = Convert.ToDecimal(listB[i].ToString());
				decimal breedte = 0;
				
				decimal Tgewicht = Convert.ToDecimal(listH[i].ToString());


				
				ScriptRecordset rsItem = this.GetRecordset("R_ITEM", "CODE, FK_ITEMUNIT, FK_ITEMGROUP", string.Format("CODE = '{0}'", ItemCode), "");
				rsItem.MoveFirst();
				
				

				if (rsItem != null && rsItem.RecordCount == 0)
				{
					MessageBox.Show("Geen overeenkomstig artikel kunnen vinden. Artikel: " + ItemCode);
				}
				else if (aantal == 0)
				{
					MessageBox.Show("Artikel: " + ItemCode + " heeft geen aantal ingevuld.");
				}
				
				else
				{					
					decimal type = Convert.ToDecimal(rsItem.Fields["FK_ITEMUNIT"].Value.ToString());
					decimal AGroup = Convert.ToDecimal(rsItem.Fields["FK_ITEMGROUP"].Value.ToString());

					// Artikleeenheden Plaat en Rooster, lengte en breedte
					if (type == 10 || type == 15 || type == 30)
					{						
						lengte = lengte / 1000;
						breedte = breedte / 1000;
					}

					// Artikleeenheden met een lengte maat
					else if (type == 11 || type == 17 || type == 20 || type == 23 || type == 24 || type == 31 || type == 32)
					{
						lengte = lengte / 1000;
						breedte = 0;
					}

					// Artikleeenheid Trapboom
					else if (type == 22 || type == 34)
					{
						lengte = lengte;
						breedte = 0;
					}

					// Artikleenheden welke nog niet gebruikt zijn
					else
					{
						lengte = 0;
						breedte = 0;
					}


					ScriptRecordset rsJoborderItem = this.GetRecordset("R_JOBORDERDETAILITEM", "", "PK_R_JOBORDERDETAILITEM= -1", "");
					rsJoborderItem.UseDataChanges = true;
					rsJoborderItem.AddNew();



					rsJoborderItem.Fields["FK_JOBORDER"].Value = bonId;
					rsJoborderItem.Fields["FK_ITEM"].Value = rsItem.Fields["PK_R_ITEM"].Value;
					rsJoborderItem.Fields["QUANTITY"].Value = aantal;
					
					rsJoborderItem.Fields["CAMPARAMETER"].Value = merk;
					rsJoborderItem.Fields["LENGTH"].Value = Convert.ToDouble(lengte);
					rsJoborderItem.Fields["WIDTH"].Value = Convert.ToDouble(breedte);
				//	rsJoborderItem.Fields["CAMGEOMETRY"].Value = Convert.ToString(myStrValues[5]);
					
					rsJoborderItem.Update();
					
					
					
					
					
					
					
					
				
				
				
				
				}

				ListGood.Add(listD[i].ToString());

			}
		}





		MessageBox.Show("Error regels= "			+ListError.Count.ToString());
		MessageBox.Show("Overgeslagen regels= "		+ListSkip.Count.ToString());
		MessageBox.Show("Goede regels= "			+ListGood.Count.ToString());
		

		ErrorBuilder(ref SalesOrder, ref Filelocation, ref ErrorLocation);
		ErrorLog(ref ErrorLocation,ref ListError);

	}









	public void MapBuilder(ref string SalesOrder, ref string Filelocation)
	{
		string BaseFolder = @"T:\Projecten\";

		string OrderStart = SalesOrder.Substring(0, 2);

		string OrderGroup = OrderStart + "00-" + OrderStart + @"99\";

		string rootFolder = BaseFolder + OrderGroup;

		string partialFolderName = SalesOrder; // Replace with the first 5 characters you know.

		string fullPath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (fullPath != null)
		{
			Filelocation = fullPath + @"\Lijsten";
		//	MessageBox.Show(Filelocation);

			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOrder);

		}
	}

	static string FindFolder(string rootFolder, string partialFolderName, string Filelocation)
	{
		try
		{
			// Get all folders in the root directory that start with the specified prefix.
			List<string> matchingFolders = Directory.GetDirectories(rootFolder)
				.Where(folder => Path.GetFileName(folder).StartsWith(partialFolderName, StringComparison.OrdinalIgnoreCase))
				.ToList();

			if (matchingFolders.Count == 1)
			{
				return matchingFolders.First(); // Return the full path of the matching folder.
			}
			else if (matchingFolders.Count > 1)
			{
				// Handle the case where there are multiple matching folders with the same prefix.
				Console.WriteLine("Multiple folders with the same prefix found. Handle this scenario as needed.");

				ShowInputDialog2(ref matchingFolders, ref Filelocation);
				return Filelocation;
				
				
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

		return null; // Return null if no matching folder is found.
	}

	public void FileBuilder(ref string Filelocation, ref string ImportFile)
	{
		string FileExtension = @".csv";


		ImportFile = FindFiles(Filelocation, FileExtension);

		if (ImportFile != null)
		{
			ImportFile = ImportFile;
		//	MessageBox.Show(ImportFile);

			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen bestanden gevonden op: " + Filelocation);

		}
	}

	static string FindFiles(string Filelocation, string FileExtension)
	{
		string ImportFile = "";
		try
		{
			// Get all files in the directory that ends with the specified suffix.
			List<string> matchingFiles = Directory.GetFiles(Filelocation)
				.Where(file => Path.GetFileName(file).EndsWith(FileExtension, StringComparison.OrdinalIgnoreCase))
				.ToList();
			
		
			if (matchingFiles.Count == 1)
			{
				return matchingFiles.First(); // Return the full path of the matching folder.
			}
			else if (matchingFiles.Count > 1)
			{
				// Handle the case where there are multiple matching folders with the same prefix.

				ShowInputDialog3(ref matchingFiles, ref ImportFile);

				return ImportFile;

			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

		return null; // Return null if no matching folder is found.
	}

	public void ErrorBuilder(ref string SalesOrder, ref string Filelocation, ref string ErrorLocation)
	{
		string BaseFolder = @"T:\Projecten\";

		string OrderStart = SalesOrder.Substring(0, 2);

		string OrderGroup = OrderStart + "00-" + OrderStart + @"99\";

		string rootFolder = BaseFolder + OrderGroup;

		string partialFolderName = SalesOrder; // Replace with the first 5 characters you know.

		string fullPath = FindFolder(rootFolder, partialFolderName, Filelocation);

		if (fullPath != null)
		{
			ErrorLocation = fullPath + @"\ALM-Errors";
			//	MessageBox.Show(Filelocation);

			// Now you can use 'fullPath' to access the folder.
		}
		else
		{
			MessageBox.Show("Geen map gevonden op: " + rootFolder + SalesOrder);

		}
	}
	
	public void ErrorLog(ref string ErrorLocation, ref List<String> ListError)
	{
		string datum = DateTime.Now.ToString();
		string datum1 = datum.Replace(":", "_");
		
		string ErrorFile = ErrorLocation + @"\Error - (" + datum1 +@").txt";
		try
		{
			// Write each item in the list to the file
			using (StreamWriter writer = new StreamWriter(ErrorFile))
			{
				foreach (string item in ListError)
				{
					writer.WriteLine(item);
				}
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("Error: " + ex.Message.ToString());
		}

	}  //creeeren van log voor overgeslagen regels
	
	

	
	
	private static DialogResult ShowInputDialog1(ref string SalesOrder)
	{
		
		
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 1.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Tekla project nummer";
		inputBox.Controls.Add(label);

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(200, 25);
		textBox.Location = new System.Drawing.Point(100, 60);
		textBox.Text = SalesOrder;
		inputBox.Controls.Add(textBox);
		
		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);
		
		inputBox.AcceptButton = okButton;
		
		DialogResult result = inputBox.ShowDialog();
		SalesOrder = textBox.Text;
		return result;
	} // bevestigen of wijzigen van het ordernummer

	private static DialogResult ShowInputDialog2(ref List<string> matchingFolders, ref string Filelocation)
	{
		
		
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 1.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Tekla project naam";
		inputBox.Controls.Add(label);

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFolders;
		combo1.Size = new System.Drawing.Size(200, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(100, 60);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);
		
		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);

		inputBox.AcceptButton = okButton;

		DialogResult result = inputBox.ShowDialog();
		Filelocation = combo1.SelectedValue.ToString();
		return result;
	}  // juiste map kiezen als er meerdere mappen zijn welke beginnen met het ordernummer

	private static DialogResult ShowInputDialog3(ref List<string> matchingFiles, ref string ImportFile)
	{
		
		
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 1.0)";

		System.Windows.Forms.Label label = new Label();
		label.Size = new System.Drawing.Size(95, 25);
		label.Location = new System.Drawing.Point(5, 60);
		label.Text = "Import lijst";
		inputBox.Controls.Add(label);

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFiles;
		combo1.Size = new System.Drawing.Size(200, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(100, 60);
		combo1.DropDownStyle = ComboBoxStyle.DropDownList;
		inputBox.Controls.Add(combo1);

		Button okButton = new Button();
		okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
		okButton.Name = "okButton";
		okButton.Size = new System.Drawing.Size(75, 23);
		okButton.Text = "&OK";
		okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, size.Height - 40);
		inputBox.Controls.Add(okButton);

		inputBox.AcceptButton = okButton;

		DialogResult result = inputBox.ShowDialog();
		ImportFile = combo1.SelectedValue.ToString();
		return result;
	}  // juiste bestand kiezen om te importeren vanaf de gekozen map
	
}