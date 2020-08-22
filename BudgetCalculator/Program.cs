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
			var jsonString = File.ReadAllText(Directory.GetCurrentDirectory() + "configuration.json");
			var configuration = JsonConvert.DeserializeObject<Configuration>(jsonString);


			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var package = new ExcelPackage(new FileInfo("MyWorkbook.xlsx")))
			{

			}
		}
	}
}
