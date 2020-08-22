using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetCalculator
{
	public class Program
	{
		public static Configuration configuration;

		static void Main(string[] args)
		{
			var jsonString = File.ReadAllText(Directory.GetCurrentDirectory() + "\\configuration.json");
			configuration = JsonConvert.DeserializeObject<Configuration>(jsonString);
			
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var package = new ExcelPackage();
			
			string exportFile = configuration.ExportPathName + configuration.ExportFileName + ".xlsx";
			if (File.Exists(exportFile))
			{
				package = new ExcelPackage(new FileInfo(exportFile));
				//Check if the last page is the same or higher than last year value. If its not then add more pages.
			}
			else
			{
				// THIS IS WHERE THE HEADER IS MADE. Create duplicates for several years (from 2019 to 2050?)
			}
			foreach (var excelFileSearch in configuration.ExcelFilesToRead)
			{
				var fullFileName = $"{excelFileSearch}.xlsx";
				var file = new ExcelPackage();

				try
				{
					FileInfo fileInfo = new FileInfo($"{configuration.ImportPathName}{fullFileName}");
					file = new ExcelPackage(fileInfo);

					CategorizeFile(package, file);

				}
				catch(Exception e)
				{
					Console.WriteLine(e.Message);
					continue;
				}
			}
			

			Console.ReadKey();
		}

		private static void CategorizeFile(ExcelPackage package, ExcelPackage file)
		{
			ExcelWorksheet worksheet = file.Workbook.Worksheets[0];
			string value = "";
			var endRow = worksheet.Dimension.End.Row;
			for (int i = configuration.StartRow; i < endRow; i++)
			{
				value = worksheet.Cells[i, configuration.ValueColumn].Value?.ToString();
				if (string.IsNullOrEmpty(value))
					break;
				var purchaseSourceName = worksheet.Cells[i, configuration.NameColumn].Value?.ToString();
				string selectedName = "";
				foreach(var c in configuration.ExportCategories)
				{
					foreach(var name in c.PurchaseSourceNames)
					{
						if (name.Equals(purchaseSourceName))
						{
							selectedName = c.CategoryName;
							break;
						}
					}

					if (!string.IsNullOrEmpty(selectedName))
					{
						break;
					}
				}

				if (string.IsNullOrEmpty(selectedName))
				{
					selectedName = configuration.CategoryOthersName;

					if(configuration.ListUpPurchaseSourcesInOthers)
						Console.WriteLine($"Added {purchaseSourceName} to {configuration.CategoryOthersName}");
				}

				var date = Convert.ToDateTime(worksheet.Cells[i, configuration.DateColumn].Value?.ToString());
				var packageWorksheet = package.Workbook.Worksheets[date.Year.ToString()];
				var packageColumn = date.Month + 1;
				var packageLastRow = packageWorksheet.Dimension.Rows;
				for(int n = configuration.StartRow, n < packageLastRow, n++)
				{

				}
				bool newCategory = false;


				


				Console.WriteLine(value);
			}
		}
	}
}
