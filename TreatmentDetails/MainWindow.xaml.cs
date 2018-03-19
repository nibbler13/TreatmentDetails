using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TreatmentDetails {
	public partial class MainWindow : Window {
		private List<ItemFilial> filials = new List<ItemFilial>();
		private string bdate = string.Empty;
		private string fdate = string.Empty;
		private string treatcount = string.Empty;
		private string mkbcodes = string.Empty;
		private string filid = string.Empty;
		private string filname = string.Empty;
		private string[] selectedFilials = new string[0];

		public MainWindow() {
			InitializeComponent();

			BackgroundWorker backgroundWorker = new BackgroundWorker();
			backgroundWorker.DoWork += BackgroundWorkerFilials_DoWork;
			backgroundWorker.RunWorkerCompleted += BackgroundWorkerFilials_RunWorkerCompleted;
			backgroundWorker.RunWorkerAsync();

			Closed += (s, e) => { DataHandle.CloseConnection(); };

			DatePickerBegin.SelectedDate = DateTime.Now.AddDays(-30);
			DatePickerFinish.SelectedDate = DateTime.Now;
			TextBoxMkbCodes.Text = "g90.9";
			TextBoxTreatCount.Text = "30";
			ListBoxSelected.Items.Add("СУЩ");
			ListBoxSelected.Items.Add("МДМ");
			ListBoxSelected.Items.Add("М-СРЕТ");
		}

		private void BackgroundWorkerFilials_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			foreach (ItemFilial item in filials)
				ListBoxTotal.Items.Add(item.SHORTNAME);
		}

		private void BackgroundWorkerFilials_DoWork(object sender, DoWorkEventArgs e) {
			filials = DataHandle.GetFilialList();
			filials = filials.OrderBy(p => p.SHORTNAME).ToList();
		}

		private void Button_Click(object sender, RoutedEventArgs e) {
			string checkResult = string.Empty;

			if (DatePickerBegin.SelectedDate == null)
				checkResult += Environment.NewLine + "Дата начала";

			if (DatePickerFinish.SelectedDate == null)
				checkResult += Environment.NewLine + "Дата окончания";

			if (ListBoxSelected.Items.Count == 0)
				checkResult += Environment.NewLine + "Филиал";

			if (string.IsNullOrEmpty(TextBoxTreatCount.Text) ||
				string.IsNullOrWhiteSpace(TextBoxTreatCount.Text))
				checkResult += Environment.NewLine + "Количество лечений";

			if (string.IsNullOrEmpty(TextBoxMkbCodes.Text) ||
				string.IsNullOrWhiteSpace(TextBoxMkbCodes.Text))
				checkResult += Environment.NewLine + "Код МКБ-10";

			if (!string.IsNullOrEmpty(checkResult)) {
				MessageBox.Show(
					this, 
					"Не указаны следующие параметры:" + checkResult, 
					string.Empty, 
					MessageBoxButton.OK, 
					MessageBoxImage.Information);
				return;
			}
			
			IsEnabled = false;

			bdate = DatePickerBegin.SelectedDate.Value.ToShortDateString();
			fdate = DatePickerFinish.SelectedDate.Value.ToShortDateString();
			treatcount = TextBoxTreatCount.Text;
			mkbcodes = TextBoxMkbCodes.Text;
			List<string> values = new List<string>();
			foreach (string item in ListBoxSelected.Items)
				values.Add(item);
			selectedFilials = values.ToArray();
			
			BackgroundWorker backgroundWorker = new BackgroundWorker();
			backgroundWorker.WorkerReportsProgress = true;
			backgroundWorker.ProgressChanged += BackgroundWorkerMain_ProgressChanged;
			backgroundWorker.DoWork += BackgroundWorkerMain_DoWork;
			backgroundWorker.RunWorkerCompleted += BackgroundWorkerMain_RunWorkerCompleted;
			backgroundWorker.RunWorkerAsync();
		}

		private void BackgroundWorkerMain_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			if (e.Error != null) {
				MessageBox.Show(
					this, 
					e.Error.Message + Environment.NewLine + e.Error.StackTrace, 
					"Ошибка", 
					MessageBoxButton.OK, 
					MessageBoxImage.Error);
			} else {
				MessageBox.Show(this, "Завершено", "", MessageBoxButton.OK, MessageBoxImage.Information);
			}

			IsEnabled = true;
		}

		private void BackgroundWorkerMain_DoWork(object sender, DoWorkEventArgs e) {
			List<string> treatcodes = new List<string>();

			string[] mkbcodesArray = mkbcodes.Split(',');

			BackgroundWorker backgroundWorker = sender as BackgroundWorker;
			double progressCurrent = 0;
			double progressStep = 20.0 / (mkbcodesArray.Length + selectedFilials.Length);

			foreach (string mkbcode in mkbcodesArray) {
				foreach (string filial in selectedFilials) {
					backgroundWorker.ReportProgress((int)progressCurrent, "Получение данных о лечениях для диагноза " + mkbcode + " для филиала " + filial);
					filid = filials.First(i => i.SHORTNAME == filial).FILID;
					treatcodes.AddRange(DataHandle.GetTreatcodes(bdate, fdate, treatcount, mkbcode, filid));
					progressCurrent += progressStep;
				}
			}

			if (treatcodes.Count == 0) {
				MessageBox.Show("Не удалось найти лечения по заданным параметрам", "", MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			List<ItemTreatmentDetails> details = DataHandle.GetDetails(treatcodes, backgroundWorker, 20);
			if (details.Count == 0) {
				MessageBox.Show("Не удалось получить информацию про лечения", "", MessageBoxButton.OK, MessageBoxImage.Information);
				return;
			}

			string filePrefix = "Отчет по использованию направлений " + mkbcodes + " с " + bdate + " по " + fdate;
			string resultFile = NpoiExcel.WriteDataToExcel(details, filePrefix, backgroundWorker, 90);

			if (File.Exists(resultFile))
				Process.Start(resultFile);

			backgroundWorker.ReportProgress(100, "Завершено");
		}

		private void BackgroundWorkerMain_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			ProgressBarMain.Value = e.ProgressPercentage;
			TextBoxProgress.Text = e.UserState.ToString();
		}

		private void ButtonToRight_Click(object sender, RoutedEventArgs e) {
			foreach (string item in ListBoxTotal.SelectedItems) {
				if (ListBoxSelected.Items.Contains(item))
					continue;

				ListBoxSelected.Items.Add(item);
			}
		}

		private void ButtonToLeft_Click(object sender, RoutedEventArgs e) {
			List<string> items = new List<string>();
			foreach (string item in ListBoxSelected.SelectedItems)
				items.Add(item);

			foreach (string item in items)
				ListBoxSelected.Items.Remove(item);
		}

		private void ListBoxTotal_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
			ButtonToRight_Click(sender, new RoutedEventArgs());
		}

		private void ListBoxSelected_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
			ButtonToLeft_Click(sender, new RoutedEventArgs());
		}
	}
}
