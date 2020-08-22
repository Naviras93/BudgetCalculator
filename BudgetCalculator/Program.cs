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
	class Program
	{
		static void Main(string[] args)
		{
			var jsonString = File.ReadAllText(Directory.GetCurrentDirectory() + "\\configuration.json");
			var configuration = JsonConvert.DeserializeObject<Configuration>(jsonString);


			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var package = new ExcelPackage())
			{
				foreach(var excelFileSearch in configuration.ExcelFilesToRead)
				{
					var fullFileName = $"{excelFileSearch}.xlsx";
					var file = new ExcelPackage();

					try
					{
						FileInfo fileInfo = new FileInfo($"{configuration.ImportPathName}{fullFileName}");
						file = new ExcelPackage(fileInfo);
					}
					catch
					{
						Console.WriteLine($"ERROR: Could not find {fullFileName} in {configuration.ImportPathName}");
						continue;
					}

					ExcelWorksheet worksheet = file.Workbook.Worksheets[0];
					string value = "";
					int row = configuration.StartRow;
					do
					{
						value = worksheet.Cells[row, configuration.Column].Value?.ToString();
						Console.WriteLine(value);
						row++;
					}
					while (!string.IsNullOrEmpty(value));
				}
			}

			Console.ReadKey();
		}
	}
}
