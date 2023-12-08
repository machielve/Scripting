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
		string SalesOrder = "6830";
		
		ShowInputDialog1(ref SalesOrder);

		MapBuilder(ref SalesOrder, ref Filelocation);

		FileBuilder(ref Filelocation, ref ImportFile);

		MessageBox.Show(ImportFile);
		
		var reader = new StreamReader(File.OpenRead(ImportFile));



		ErrorLog(ref ErrorLocation);

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
	
	public void ErrorLog(ref string ErrorLocation)
	{

	}  //creeeren van log voor overgeslagen regels
	
	

	
	
	private static DialogResult ShowInputDialog1(ref string SalesOrder)
	{
		
		
		System.Drawing.Size size = new System.Drawing.Size(400, 400);
		Form inputBox = new Form();

		inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		inputBox.ClientSize = size;
		inputBox.Text = "Cluedo (Tekla import 1.0)";

		System.Windows.Forms.TextBox textBox = new TextBox();
		textBox.Size = new System.Drawing.Size(size.Width - 75, 23);
		textBox.Location = new System.Drawing.Point(60, 10);
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

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFolders;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
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

		System.Windows.Forms.ComboBox combo1 = new ComboBox();
		combo1.DisplayMember = "TOTAAL";
		combo1.ValueMember = "CODE";
		combo1.DataSource = matchingFiles;
		combo1.Size = new System.Drawing.Size(275, 25);
		combo1.DropDownWidth = 500;
		combo1.Location = new System.Drawing.Point(5, 150);
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