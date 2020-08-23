using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetCalculator
{
	public class Model
	{
		public string Year { get; set; }
		public int Column { get; set; }
		public string Category { get; set; }
		public decimal Amount { get; set; }
		public bool NewCategory { get; set; }
	}
}
