﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetCalculator
{
	public class Configuration
	{
		public bool IgnoreProfit { get; set; }
		public int StartRow { get; set; }
		public int ValueColumn { get; set; }
		public int NameColumn { get; set; }
		public int DateColumn { get; set; }
		public int CoverYearsBehind { get; set; }
		public List<string> MonthHeaders { get; set; }
		public string ImportPathName { get; set; }
		public List<string> ExcelFilesToRead { get; set; }
		public List<ExportCategory> ExportCategories { get; set; }
		public string CategoryOthersName { get; set; }
		public bool ListUpPurchaseSourcesInOthers { get; set; }
		public string ExportFileName { get; set; }
		public string ExportPathName { get; set; }

		public class ExportCategory
		{
			public string CategoryName { get; set; }
			public List<string> PurchaseSourceNames { get; set; }
		}
	}
}
