using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;                                                         // Обязательно добавлять через "Ссыллки" -> Обзор и выбирать библиотеку Microsoft.Office.Interop.Excel.dll в папке установки Ms VS
using System.Threading;

namespace BankSum
{
	public partial class BankSum : Form
		{
		public BankSum()
			{
			InitializeComponent();

			// Парсин кура валют с сайта ЦБ РФ:
			CultureInfo.DefaultThreadCurrentCulture = CultureInfo.GetCultureInfo("ru-RU");
			//CultureInfo ci = new System.Globalization.CultureInfo("ru-Ru");
			WebClient client = new WebClient();
			var xml = client.DownloadString("https://www.cbr.ru/scripts/XML_daily.asp?");
			XDocument xdoc = XDocument.Parse(xml);
			var el = xdoc.Element("ValCurs").Elements("Valute");
			string dollar = el.Where(x => x.Attribute("ID").Value == "R01235").Select(x => x.Element("Value").Value).FirstOrDefault();
			string eur = el.Where(x => x.Attribute("ID").Value == "R01239").Select(x => x.Element("Value").Value).FirstOrDefault();
			kursval_dollar.Text = Math.Round(Convert.ToDecimal(dollar), 2).ToString();
			kursval_euro.Text = Math.Round(Convert.ToDecimal(eur), 2).ToString();

			// Отображение текущей даты в заголовке волютного бокса GroupBox5
			Date.Text = DateTime.Now.ToString("dd.MM.yyyy");
			}
		private void BankSum_Load(object sender, EventArgs e)
			{
			// Снятие фокуа с первого верхнего TextBox при старте программы и перенос его на PictureBox - лого ROSGP
			this.ActiveControl = pictureBox1;
		}

		// Расчет количества КУПЮР в суммы (первый расчёт стандартный if, затем в вите ТЕРНАРНОЙ операции (в одну строку)):
		
		// Описываем обработчик событияй для TextBox tb10pk - событие TextGanged 
		private void tb10pk_TextChanged(object sender, EventArgs e)																														// номинал 10 рублей (купюра)
			{
			// Задаём переменную result типа boolc присваением ей целочисленного значения с преобразованием из строкового в 32х разрядное,
			// а также сокраoftv на выходе tb10pk в res
			bool result = int.TryParse(tb10pk.Text, out int res);
			// Описываем функцию if/else в тернарном виде
			lb10pk.Text = result ? (res * 10).ToString() : "0";
			// Обязательно errfrpsdftv ссылку на метод sn(); для подсчёта общей суммы купюр в реальном времени
			sn();
			}
		// Описываем обработчик события для события TextBox tb10pk - события KeyPress 
		private void tb10pk_KeyPress(object sender, KeyPressEventArgs e)
			{
			// Условие для ввода только цифр (буквы и знакки пунтуации отключены)
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b') return;
			// Вывод MessageBox с уведомление о запрете выывода символов кроме цифр -
			// данное окно уведомлений сделанно как пример, и выводится только для первого TextBox - tb10pk
			// в последующих Textbox оно упращено и запрет вводв символов кроме цифр ничем не сообщается - тишина, просто не вводятся
			else
				{
				MessageBox.Show("Ввод любых символов кроме цифр - ЗАПРЕЩЁН! Введите число!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
				e.Handled = true;
				}
			}

		private void tb50p_TextChanged(object sender, EventArgs e)                  // номинал 50 рублей
			{
			bool result = int.TryParse(tb50p.Text, out int res);
			lb50p.Text = result ? (res * 50).ToString() : "0";
			sn();
			}
		private void tb50p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b') 
					return;
			else
				{
				e.Handled = true;
				}
			}
		

		private void tb100p_TextChanged(object sender, EventArgs e)                // номинал 100 рублей
			{
			bool result = int.TryParse(tb100p.Text, out int res);
				lb100p.Text = result ? (res * 100).ToString() : "0";
			sn();
			}
		private void tb100p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}
		
		private void tb200p_TextChanged(object sender, EventArgs e)               // номинал 200 рублей
			{
			bool result = int.TryParse(tb200p.Text, out int res);
			lb200p.Text = result ? (res * 200).ToString() : "0";
			sn();
			}
		private void tb200p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;

			else
				{
				e.Handled = true;
				}
			}
		
		private void tb500p_TextChanged(object sender, EventArgs e)               // номинал 500 рублей
			{
			bool result = int.TryParse(tb500p.Text, out int res);
			lb500p.Text = result ? (res * 500).ToString() : "0";
			sn();
			}
		private void tb500p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}

		private void tb1000p_TextChanged(object sender, EventArgs e)               // номинал 1000 рублей
			{
			bool result = int.TryParse(tb1000p.Text, out int res);
			lb1000p.Text = result ? (res * 1000).ToString() : "0"; ;
			sn();
			}
		private void tb1000p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}

		private void tb2000p_TextChanged(object sender, EventArgs e)               // номинал 2000 рублей
			{
			bool result = int.TryParse(tb2000p.Text, out int res);
				lb2000p.Text = result ? (res * 2000).ToString() : "0";
			sn();
			}
		private void tb2000p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}
	
		private void tb5000p_TextChanged(object sender, EventArgs e)               // номинал 5000 рублей
			{
			bool result = int.TryParse(tb5000p.Text, out int res);
			lb5000p.Text = result ? (res * 5000).ToString() : "0";
			sn();
			}
		private void tb5000p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}

			}

	// Общую сумму КУПЮР подсчитываем через метод sn() - орбязательно добавляя на него ссылку в каждый обработчик событие TexBox купюр, иначе общая сумма купюр считаться не будет!;
	public void sn()
			{
			int sn10 = Convert.ToInt32(lb10pk.Text);
			int sn50 = Convert.ToInt32(lb50p.Text);
			int sn100 = Convert.ToInt32(lb100p.Text);
			int sn200 = Convert.ToInt32(lb200p.Text);
			int sn500 = Convert.ToInt32(lb500p.Text);
			int sn1000 = Convert.ToInt32(lb1000p.Text);
			int sn2000 = Convert.ToInt32(lb2000p.Text);
			int sn5000 = Convert.ToInt32(lb5000p.Text);
			int sumk = (sn10 + sn50 + sn100 + sn200 + sn500 + sn1000 + sn2000 + sn5000);
			lbsum_k.Text = Convert.ToString(sumk);
			snm();
			}

		// Расчет количества МОНЕТ в суммы:
		private void tb05p_TextChanged(object sender, EventArgs e)               // номинал 0,5 копеек
			{
			bool result = int.TryParse(tb05p.Text, out int res);
			lb05p.Text = result ? (res * 0.5).ToString() : "0";
			sm();
			}
		private void tb05p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}

		private void tb1p_textChanged(object sendet, EventArgs e)               // номинал 1 рубль
			{
			bool result = int.TryParse(tb1p.Text, out int res);
			lb1p.Text = result ? (res * 1).ToString() : "0";
			sm();
			}
		private void tb1p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
			}
		}

		private void tb2p_TextChanged(object sender, EventArgs e)               // номинал 2 рубля
			{
			bool result = int.TryParse(tb2p.Text, out int res);
			lb2p.Text = result ? (res * 2).ToString() : "0";
			sm();
			}
		private void tb2p_KeyPress(object sendet, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}
		
		private void tb5p_TextChanged(object sender, EventArgs e)               // номинал 5 рублей
			{
			bool result = int.TryParse(tb5p.Text, out int res);
			lb5p.Text = result ? (res * 5).ToString() : "0";
			sm();
			}
		private void tb5p_KeyPress(object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true; ;
				}
			}
		

		private void tb10pm_TextChanged(object sender, EventArgs e)               // номинал 10 рублей
			{
			bool result = int.TryParse(tb10pm.Text, out int res);
			lb10pm.Text = result ? (res * 10).ToString() : "0";
			sm();
			}
		private void tb10pm_KeyPress( object sender, KeyPressEventArgs e)
			{
			if (char.IsNumber(e.KeyChar) | e.KeyChar == '\b')
				return;
			else
				{
				e.Handled = true;
				}
			}

	
		
		// Общую сумму МОНЕТ подсчитываем через метод sm();
		public void sm()
			{
			decimal sm05 = Convert.ToDecimal(lb05p.Text);
			int sm1 = Convert.ToInt32(lb1p.Text);
			int sm2 = Convert.ToInt32(lb2p.Text);
			int sm5 = Convert.ToInt32(lb5p.Text);
			int sm10 = Convert.ToInt32(lb10pm.Text);
			int summ = (sm1 + sm2 + sm5 + sm10);
			lbsum_m.Text = (sm05 + summ).ToString();
			snm();
			}

		// Итоговую сумму КУПЮР и МОНЕТ подсчитываем через метод snm(); 
		public void snm()
			{
			int sum_k = Convert.ToInt32(lbsum_k.Text);
			decimal sum_m = Convert.ToDecimal(lbsum_m.Text);
			lbsum_end.Text = Convert.ToDecimal(sum_k + sum_m).ToString();
			lbsum_end.ForeColor = Color.Blue;

			// Выводим итоговую сумму в долларах и евро по текущему ккурсу
			decimal rub = Convert.ToDecimal(lbsum_end.Text);
			decimal doll = Convert.ToDecimal(kursval_dollar.Text);
			decimal euro = Convert.ToDecimal(kursval_euro.Text);
			lb_sum_dollar.Text = Math.Round((rub / doll), 2).ToString();
			lb_sum_euro.Text = Math.Round((rub / euro), 2).ToString();
			}

		// Кнопка СОХРАНИТЬ В EXCEL как
		private void bt_SaveToExcel_Click(object sender, EventArgs e)
			{
			SaveFileDialog saveFile = new SaveFileDialog
				{
				Filter = "Excel (*.xls)|*.xls|All files (*.*)|*.*",
				Title = "Сохранить"
				};

			if (saveFile.ShowDialog() == DialogResult.OK)
				{
				using (FileStream fileExcel = new FileStream(saveFile.FileName, FileMode.Append))
				using (StreamWriter writer = new StreamWriter(fileExcel, Encoding.UTF8))
					writer.WriteLine(lbsum_end.Text);
				}
			}

		// Кнопка СОХРАНИТЬ В ОТКРЫТЫЙ ФАЙЛ EXCEL
		private void bt_SaveToOpenFileExcel_Click(object sender, EventArgs e)
			{
			Excel.Application ExcelApp = new Excel.Application();
			Excel.Workbook ExcelWorkBook;
			Excel.Worksheet ExcelWorkSheet;
			//Открываем книгу.
			ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
			//Создаем таблицу.
			ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

			ExcelApp.Visible = true;

			// Этот пример использует одну таблицу.
			Excel._Worksheet workSheet = ExcelApp.ActiveSheet;

			workSheet.Cells[1, "A"] = lbsum_end.Text; //указывает куда сохранять данные из label
																	//workSheet.Cells[1, "B"] = textbox.text; //создаем 2 столбец
			}

		// Кнопка СБРОС всех значений
		public void bt_Clear_Click(object sender, EventArgs e)
			{
			_ = tb10pk.Text != "" ? tb10pk.Text = tb10pk.Text.Remove(0) : tb10pk.Text = string.Empty;	lb10pk.Text = "0";	ActiveControl = tb10pk;
			_ = tb50p.Text != "" ? tb50p.Text = tb50p.Text.Remove(0) : tb50p.Text = string.Empty;	lb50p.Text = "0";	ActiveControl = tb50p;
			_ = tb100p.Text != "" ? tb100p.Text = tb100p.Text.Remove(0) : tb100p.Text = string.Empty;	lb100p.Text = "0";	ActiveControl = tb100p;
			_ = tb200p.Text != "" ? tb200p.Text = tb200p.Text.Remove(0) : tb200p.Text = string.Empty;	lb200p.Text = "0";	ActiveControl = tb200p;
			_ = tb500p.Text != "" ? tb500p.Text = tb500p.Text.Remove(0) : tb500p.Text = string.Empty;	lb500p.Text = "0";	ActiveControl = tb500p;
			_ = tb1000p.Text != "" ? tb1000p.Text = tb1000p.Text.Remove(0) : tb1000p.Text = string.Empty;	lb1000p.Text = "0";	ActiveControl = tb1000p;
			_ = tb2000p.Text != "" ? tb2000p.Text = tb2000p.Text.Remove(0) : tb2000p.Text = string.Empty;	lb2000p.Text = "0";	ActiveControl = tb2000p;
			_ = tb5000p.Text != "" ? tb5000p.Text = tb5000p.Text.Remove(0) : tb5000p.Text = string.Empty;	lb5000p.Text = "0";	ActiveControl = tb5000p;
			_ = tb05p.Text != "" ? tb05p.Text = tb05p.Text.Remove(0) : tb05p.Text = string.Empty;	lb05p.Text = "0";	ActiveControl = tb05p;
			_ = tb1p.Text != "" ? tb1p.Text = tb1p.Text.Remove(0) : tb1p.Text = string.Empty;	lb1p.Text = "0";	ActiveControl = tb1p;
			_ = tb2p.Text != "" ? tb2p.Text = tb2p.Text.Remove(0) : tb2p.Text = string.Empty;	lb2p.Text = "0";	ActiveControl = tb2p;
			_ = tb5p.Text != "" ? tb5p.Text = tb5p.Text.Remove(0) : tb5p.Text = string.Empty;	lb5p.Text = "0";	ActiveControl = tb5p;
			_ = tb10pm.Text != "" ? tb10pm.Text = tb10pm.Text.Remove(0) : tb10pm.Text = string.Empty;	lb10pm.Text = "0";	ActiveControl = tb10pm;
			_ = tb_trash.Text != "" ? tb_trash.Text = tb_trash.Text.Remove(0) : tb_trash.Text = string.Empty; lb10pm.Text = "0"; ActiveControl = tb_trash;

			lbsum_m.Text = "0,00";
			lbsum_end.Text = "0,00";
			lb_sum_dollar.Text = "0,00";
			lb_sum_euro.Text = "0,00";
			ActiveControl = pictureBox1;
			}
		
		private void pictureBox1_Click_1(object sender, EventArgs e)
			{
			MessageBox.Show("Вы сейчас будете перенаправлены на YouTube канал создателя данной программы! Не забудьте поставить лайки и подписаться на канал! Большое спасибо!", "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Process.Start("https://www.youtube.com/c/RosGamePlay");           // Process - даёт возможность добавлять гипперссылки
			}

		private void pictureBox1_MouseEnter(object sender, EventArgs e)
			{
			ToolTip t = new ToolTip();
			t.SetToolTip(pictureBox1, "YouTube канал Автора");
			}

		
		}
}
