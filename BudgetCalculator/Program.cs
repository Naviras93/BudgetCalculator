using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetCalculator
{
	public class Program
	{
		public static Configuration configuration;
		public static string exportFile;
		public static bool newFile;

		static void Main(string[] args)
		{
			var jsonString = File.ReadAllText(Directory.GetCurrentDirectory() + "\\configuration.json", System.Text.Encoding.GetEncoding(1252));
			configuration = JsonConvert.DeserializeObject<Configuration>(jsonString);
			
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var package = new ExcelPackage();
			
			exportFile = configuration.ExportPathName + configuration.ExportFileName + ".xlsx";
			var earliestYear = DateTime.Now.Year + configuration.CoverYearsBehind;
			if (File.Exists(exportFile))
			{
				newFile = false;
				package = new ExcelPackage(new FileInfo(exportFile));
				if(int.Parse(package.Workbook.Worksheets.Last().Name) > earliestYear)
				{
					AddPages(package, earliestYear);
				}
			}
			else
			{
				newFile = true;
				AddPages(package, earliestYear);
			}
			List<Model> models = new List<Model>();
			foreach (var excelFileSearch in configuration.ExcelFilesToRead)
			{
				var fullFileName = $"{excelFileSearch}.xlsx";
				var file = new ExcelPackage();

				try
				{
					FileInfo fileInfo = new FileInfo($"{configuration.ImportPathName}{fullFileName}");
					file = new ExcelPackage(fileInfo);

					FillModel(file, models);

				}
				catch(Exception e)
				{
					Console.WriteLine(e.Message);
					continue;
				}
			}

			CreateFile(package, models);
			if (newFile)
				Console.WriteLine($"Created {configuration.ExportFileName}.xlsx");
			else
				Console.WriteLine($"Updated {configuration.ExportFileName}.xlsx");

			Console.ReadKey();
		}

		private static void AddPages(ExcelPackage package, int untilYear)
		{
			for (int i = DateTime.Now.Year; i <= untilYear; i++)
			{
				if(package.Workbook.Worksheets[i.ToString()] == null)
				{
					var sheet = package.Workbook.Worksheets.Add(i.ToString());
					DesignPage(sheet);
				}
			}
		}

		private static void DesignPage(ExcelWorksheet sheet)
		{
			int firstColumn = 2;
			int lastColumn = firstColumn;
			foreach (var header in configuration.Headers)
			{
				sheet.Cells[1, lastColumn].Value = header;
				lastColumn++;
			}
			if (lastColumn > firstColumn)
			{
				var headerCells = sheet.SelectedRange[1, firstColumn, 1, lastColumn - 1];
				headerCells.Style.Font.Bold = true;
				headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
				headerCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
			}
		}

		private static void FillModel(ExcelPackage file, List<Model> models)
		{
			CultureInfo culture = new CultureInfo("sv-SE");
			ExcelWorksheet worksheet = file.Workbook.Worksheets[0];
			string value = "";
			var endRow = worksheet.Dimension.End.Row;
			for (int i = configuration.StartRow; i < endRow; i++)
			{
				value = worksheet.Cells[i, configuration.ValueColumn].Value?.ToString();
				if (string.IsNullOrEmpty(value))
					break;
				if (configuration.IgnoreProfit && !value.StartsWith("-"))
					continue;
				var purchaseSourceName = worksheet.Cells[i, configuration.NameColumn].Value?.ToString();
				string selectedName = "";
				foreach(var c in configuration.ExportCategories)
				{
					selectedName = "";
					foreach (var name in c.PurchaseSourceNames)
					{
						//if (purchaseSourceName.Contains(name))
						//{
						//	selectedName = c.CategoryName;
						//	break;
						//}
						if(culture.CompareInfo.IndexOf(purchaseSourceName, name, CompareOptions.IgnoreCase) >= 0)
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

				var date = Convert.ToDateTime(worksheet.Cells[i, configuration.DateColumn].Text);
				var dateYear = date.Year.ToString();
				var packageColumn = date.Month;
				var model = models.Where(x => x.Year.Equals(dateYear) && x.Column == packageColumn && x.Category.Equals(selectedName)).FirstOrDefault();
				if(model == null)
				{
					model = new Model
					{
						Category = selectedName,
						Column = packageColumn,
						Year = dateYear,
						Amount = decimal.Parse(value.Replace(".", "")),
					};
					models.Add(model);
				}
				else
					model.Amount += decimal.Parse(value.Replace(".", ""));
			}
		}

		private static void CreateFile(ExcelPackage package, List<Model> models)
		{
			foreach(var model in models)
			{
				bool newCategory = true;
				var excelWorksheet = package.Workbook.Worksheets[model.Year];
				int lastRow = 2;
				if (excelWorksheet == null)
				{
					excelWorksheet = package.Workbook.Worksheets.Add(model.Year);
				}
				else
				{
					lastRow = excelWorksheet.Dimension.Rows;
				}

				for (int i = 2; i <= lastRow; i++)
				{
					if (excelWorksheet.Cells[i, 1].Value.Equals(model.Category))
					{
						excelWorksheet.Cells[i, model.Column].Value = model.Amount;
						newCategory = false;
						break;
					}
				}
				if (newCategory)
				{
					lastRow++;
					excelWorksheet.Cells[lastRow, 1].Value = model.Category;
					excelWorksheet.Cells[lastRow, model.Column].Value = model.Amount;
					excelWorksheet.Cells.AutoFitColumns();
				}
			}

			if (newFile)
				package.SaveAs(new FileInfo(exportFile));
			else
				package.Save();
		}

	}
}
