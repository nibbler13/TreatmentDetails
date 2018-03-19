using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace TreatmentDetails {
	class NpoiExcel {
		public static string WriteDataToExcel(List<ItemTreatmentDetails> details, string resultFilePrefix, BackgroundWorker backgroundWorker, double progressCurrent) {
			string templateFile = SystemLogging.AssemblyDirectory + "Template.xlsx";
			double progressStep = 10.0 / details.Count;
			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			if (!File.Exists(templateFile))
				return "Не удалось найти файл шаблона: " + templateFile;

			string resultPath = Path.Combine(SystemLogging.AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			string resultFile = Path.Combine(resultPath, resultFilePrefix + ".xlsx");
			
			IWorkbook workbook;
			using (FileStream stream = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
				workbook = new XSSFWorkbook(stream);

			int rowNumber = 1;
			int columnNumber = 0;

			ISheet sheet = workbook.GetSheet("Data");

			foreach (ItemTreatmentDetails item in details) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Запись в excel строки " + rowNumber + " / " + details.Count);
				IRow row = sheet.CreateRow(rowNumber);
				//Console.WriteLine("create row: " + row);

				DateTime treatDate = DateTime.ParseExact(item.TREATDATE, "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture);
				double period = -1;
				string referralPeriodMax = string.Empty;

				foreach (ItemReferral referral in item.Referrals) {
					if (string.IsNullOrEmpty(referral.TREATDATE1))
						continue;

					DateTime refDate = DateTime.ParseExact(referral.TREATDATE1, "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture);
					double days = (refDate - treatDate).TotalDays;
					if (days > period)
						period = days;
				}

				if (period > -1)
					referralPeriodMax = period.ToString();


				string[] array = new string[] {
					item.TREATDATE,
					item.FILIALNAME,
					item.DEPNAME,
					item.DOCNAME,
					item.PATIENTNAME,
					item.HISTNUM,
					item.BDATE,
					item.MKBCODE,
					item.Referrals.Count.ToString(),
					referralPeriodMax
				};

				//Console.WriteLine("values: " + string.Join(", ", array));

				foreach (string value in array) {
					try {
						ICell cell = row.CreateCell(columnNumber);
						Console.WriteLine("create column: " + columnNumber);
						string valueToWrite = value.Replace(" 0:00:00", "");

						if (double.TryParse(valueToWrite, out double result))
							cell.SetCellValue(result);
						else
							cell.SetCellValue(valueToWrite);
					} catch (Exception e) {
						Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
					}

					columnNumber++;
				}

				foreach (ItemReferral referral in item.Referrals) {
					string name = referral.SCHNAME;
					int isActivated = string.IsNullOrEmpty(referral.TREATDATE1) ? 0 : 1;

					for (int i = columnNumber; i < 16000; i++) {
						string currentValue = "";
						try {
							if (i < sheet.GetRow(0).LastCellNum)
								currentValue = sheet.GetRow(0).GetCell(i).StringCellValue;
						} catch (Exception) {
						}

						try {
							if (string.IsNullOrEmpty(currentValue)) {
								ICell cellHeader = sheet.GetRow(0).CreateCell(i);
								cellHeader.SetCellValue(name);

								ICell cellValue = row.CreateCell(i);
								cellValue.SetCellValue(isActivated);
								break;
							}
						} catch (Exception e) {
							Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
						}

						try {
							if (currentValue.Equals(name)) {
								ICell cell = row.CreateCell(i);
								cell.SetCellValue(isActivated);
								break;
							}
						} catch (Exception e) {
							Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
						}
					}
				}

				columnNumber = 0;
				rowNumber++;
				progressCurrent += progressStep;
			}

			sheet = workbook.GetSheet("Services");
			int rowUsed = 1;
			Console.WriteLine("=======================================");
			Console.WriteLine("sheet.LastRowNum: " + sheet.LastRowNum);
			foreach (ItemTreatmentDetails item in details) {
				foreach (ItemReferral referral in item.Referrals) {
					string name = referral.SCHNAME;
					bool isPresent = false;
					bool isActivated = (string.IsNullOrEmpty(referral.TREATDATE1)) ? false : true;
					double.TryParse(referral.SCOUNT, out double currentTotal);
					double.TryParse(referral.SCHCOUNT, out double currentActivated);

					Console.WriteLine("name: " + name + " | " + currentTotal + " " + currentActivated);

					for (int row = 1; row < rowUsed; row++) {
						try {
							string rowValue = sheet.GetRow(row).GetCell(0).StringCellValue;
							if (rowValue.Equals(name)) {
								double total = sheet.GetRow(row).GetCell(1).NumericCellValue;
								double activated = sheet.GetRow(row).GetCell(2).NumericCellValue;
								Console.WriteLine("finded: " + total + " - " + activated);

								total += currentTotal;
								activated += currentActivated;
								
								sheet.GetRow(row).GetCell(1).SetCellValue(total);
								sheet.GetRow(row).GetCell(2).SetCellValue(activated);
								isPresent = true;
								break;
							}
						} catch (Exception) {
							break;
						}
					}

					if (isPresent)
						continue;

					IRow rowLine = sheet.CreateRow(rowUsed);
					ICell cellName = rowLine.CreateCell(0);
					ICell cellTotal = rowLine.CreateCell(1);
					ICell cellActivated = rowLine.CreateCell(2);

					cellName.SetCellValue(name);
					cellTotal.SetCellValue(currentTotal);
					cellActivated.SetCellValue(currentActivated);

					rowUsed++;
				}
			}
			Console.WriteLine("=======================================");

			using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
				workbook.Write(stream);

			workbook.Close();

			Excel.Application xlApp = new Excel.Application();

			if (xlApp == null)
				return "Не удалось открыть приложение Excel";

			xlApp.Visible = false;

			Excel.Workbook wb = xlApp.Workbooks.Open(resultFile);

			if (wb == null)
				return "Не удалось открыть книгу " + resultFile;

			Excel.Worksheet ws = wb.Sheets["Data"];

			if (ws == null)
				return "Не удалось открыть лист Data";

			try {
				PerformSheet(wb, ws, xlApp);
			} catch (Exception e) {
				SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			ws = wb.Sheets["Services"];
			ws.Activate();

			try {
				PerformSheetServices(wb, ws, xlApp);
			} catch (Exception e) {
				SystemLogging.LogMessageToFile(e.Message + Environment.NewLine + e.StackTrace);
			}

			wb.Save();
			wb.Close();

			xlApp.Quit();

			return resultFile;
		}

		private static void PerformSheetServices(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			int rowUsed = ws.UsedRange.Rows.Count;
			ws.Range["D2"].Select();
			xlApp.ActiveCell.FormulaR1C1 = "=(RC[-2]-RC[-1])/RC[-2]";
			xlApp.Selection.NumberFormat = "0%";
			xlApp.Selection.AutoFill(ws.Range["D2:D" + rowUsed]);
			xlApp.Selection.AutoFilter();
			ws.AutoFilter.Sort.SortFields.Add(ws.Range["B1:B" + rowUsed], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending, Excel.XlSortDataOption.xlSortNormal);
			ws.AutoFilter.Sort.Header = Excel.XlYesNoGuess.xlYes;
			ws.AutoFilter.Sort.Apply();
		}

		private static void PerformSheet(Excel.Workbook wb, Excel.Worksheet ws, Excel.Application xlApp) {
			int columnUsed = ws.UsedRange.Columns.Count;
			ws.Columns["A:H"].Select();
			xlApp.Selection.Columns.AutoFit();
			ws.Columns["I:" + GetExcelColumnName(columnUsed)].Select();
			xlApp.Selection.ColumnWidth = 5;
			ws.Columns["C:E"].Select();
			xlApp.Selection.ColumnWidth = 20;
			ws.UsedRange.Font.Size = 8;
			ws.Rows[1].WrapText = true;
			ws.Rows[1].VerticalAlignment = Excel.Constants.xlTop;
			ws.Rows[1].RowHeight = 80;
			ws.Cells[1, 1].Select();
			xlApp.ActiveWindow.SplitRow = 1;
			xlApp.ActiveWindow.FreezePanes = true;
		}

		private static string GetExcelColumnName(int columnNumber) {
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0) {
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
	}
}
