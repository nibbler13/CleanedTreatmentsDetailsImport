using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CleanedTreatmentsDetailsImport {
	public class ItemProfitAndLoss {
		public String ObjectName { get; set; }
		public int PeriodYear { get; set; }
		public String PeriodType { get; set; }
		public String GroupNameLevel1 { get; set; }
		public String GroupNameLevel2 { get; set; }
		public String GroupNameLevel3 { get; set; }
		public double Value { get; set; }
		public int GroupSortingOrder { get; set; }
		public int ObjectSrotingOrder { get; set; }
		public int? Quarter { get; set; }
		public bool? HasData { get; set; }
	}
}
