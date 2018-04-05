using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TreatmentDetails {
	public class ItemTreatmentDetails {
		public string FILIALNAME { get; set; } = string.Empty;
		public string TREATCODE { get; set; } = string.Empty;
		public string TREATDATE { get; set; } = string.Empty;
		public string DOCNAME { get; set; } = string.Empty;
		public string DEPNAME { get; set; } = string.Empty;
		public string PATIENTNAME { get; set; } = string.Empty;
		public string HISTNUM { get; set; } = string.Empty;
		public string BDATE { get; set; } = string.Empty;
		public string MKBCODE { get; set; } = string.Empty;
		public string TREAT_TYPE { get; set; } = string.Empty;
		public List<ItemReferral> Referrals { get; set; } = new List<ItemReferral>();
	}
}
