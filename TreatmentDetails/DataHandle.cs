using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TreatmentDetails {
    class DataHandle {
		private static string sqlQuerySelectFilials = Properties.Settings.Default.MisDbSqlGetFilials;
		private static string sqlQuerySelectTreatments = Properties.Settings.Default.MisDbSqlGetTreatByMkb;
		private static string sqlQuerySelectDetails = Properties.Settings.Default.MisDbSqlGetTreatDetails;

		private static FirebirdClient firebirdClient = new FirebirdClient(
			Properties.Settings.Default.MisDbAddress,
			Properties.Settings.Default.MisDbName,
			Properties.Settings.Default.MisDbUser,
			Properties.Settings.Default.MisDbPassword);

		public static void CloseConnection() {
			firebirdClient.Close();
		}

		public static List<ItemFilial> GetFilialList() {
			List<ItemFilial> list = new List<ItemFilial>();
			
			DataTable dataTable = firebirdClient.GetDataTable(sqlQuerySelectFilials, new Dictionary<string, string>());
			if (dataTable.Rows.Count == 0)
				return list;

			foreach (DataRow row in dataTable.Rows) {
				try {
					list.Add(new ItemFilial() {
						FILID = row["FILID"].ToString(),
						FULLNAME = row["FULLNAME"].ToString(),
						SHORTNAME = row["SHORTNAME"].ToString()
					});
				} catch (Exception e) {
					MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Ошибка обработки данных",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}

			return list;
		}

		public static List<string> GetTreatcodes(string bdate, string fdate, string treatcount, string mkbcode, string filid) {
			List<string> list = new List<string>();

			string query = sqlQuerySelectTreatments.
				Replace("@bdate", bdate).
				Replace("@fdate", fdate).
				Replace("@filid", filid).
				Replace("@mkbcode", mkbcode).
				Replace("@treatcount", treatcount);
			DataTable dataTable = firebirdClient.GetDataTable(query, new Dictionary<string, string>());

			if (dataTable.Rows.Count == 0) 
				return list;

			foreach (DataRow row in dataTable.Rows) {
				try {
					list.Add(row[0].ToString());
				} catch (Exception e) {
					MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Ошибка обработки данных",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
			
			return list;
		}

		public static List<ItemTreatmentDetails> GetDetails(List<string> treatcodes, BackgroundWorker backgroundWorker, double progressCurrent) {
			List<ItemTreatmentDetails> list = new List<ItemTreatmentDetails>();

			double progressStep = 70.0 / treatcodes.Count;
			int currentTreat = 1;
			foreach (string treatcode in treatcodes) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Получение данных о лечении " + currentTreat + " / " + treatcodes.Count);
				currentTreat++;
				DataTable dataTable = firebirdClient.GetDataTable(sqlQuerySelectDetails, new Dictionary<string, string> {
					{ "@treatcode", treatcode } });
				progressCurrent += progressStep;

				if (dataTable.Rows.Count == 0)
					continue;

				ItemTreatmentDetails details = null;

				foreach (DataRow row in dataTable.Rows) {
					try {
						if (details == null) {
							details = new ItemTreatmentDetails() {
								FILIALNAME = row["FILIAL"].ToString(),
								TREATCODE = row["TREATCODE"].ToString(),
								TREATDATE = row["TREATDATE"].ToString(),
								DOCNAME = row["DOCTORNAME"].ToString(),
								DEPNAME = row["DEPNAME"].ToString(),
								PATIENTNAME = row["CLIENTNAME"].ToString(),
								HISTNUM = row["HISTNUM"].ToString(),
								BDATE = row["BDATE"].ToString(),
								MKBCODE = row["MKBCODE"].ToString(),
							};

							string listAllServices = row["LISTALLSERVICES"].ToString().ToUpper();
							details.TREAT_TYPE = listAllServices.Contains("ПЕРВИЧНЫЙ") ?
								"Первичный" : "Повторный";
						}

						string refid = row["REFID"].ToString();
						if (string.IsNullOrEmpty(refid) || string.IsNullOrWhiteSpace(refid))
							continue;

						ItemReferral refferal = new ItemReferral() {
							REFID = refid,
							SCHNAME = row["SCHNAME"].ToString(),
							SCOUNT = row["SCOUNT"].ToString(),
							SHORTNAME1 = row["REFFILIAL"].ToString(),
							TREATDATE1 = row["REFTREATDATE"].ToString(),
							FULLNAME2 = row["REFDOCTORNAME"].ToString(),
							DEPNAME1 = row["REFDEPNAME"].ToString(),
							SCHNAME1 = row["REFSCHNAME"].ToString(),
							SCHCOUNT = row["SCHCOUNT"].ToString()
						};

						details.Referrals.Add(refferal);
					} catch (Exception e) {
						MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Ошибка обработки данных",
							MessageBoxButton.OK, MessageBoxImage.Error);
					}
				}

				if (details != null)
					list.Add(details);
			}

			return list;
		}
	}
}
