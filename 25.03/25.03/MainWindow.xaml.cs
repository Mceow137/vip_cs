using System;
using System.Collections.Generic;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace _25._03
{

    public partial class MainWindow : Window
    {
        Excel.Worksheet xlSht;
        public List<string> dataList = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
        }

        public void Button_Dobavit_Click(object sender, RoutedEventArgs e)
        {
            string chifra = Text.Text;
            if (!string.IsNullOrEmpty(chifra))
            {
                dataList.Add(chifra);
                List.Items.Add(chifra);
                Text.Clear();
            }
            else
            {
                MessageBox.Show("Введите данные");
            }
        }

        public void Button_Soxranit_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application app = new Excel.Application
            {
                Visible = true,
                SheetsInNewWorkbook = 2
            };
            Excel.Workbook xlWB;
            app.DisplayAlerts = false;

            xlWB = app.Workbooks.Open(@"C:\22П-1\Жуйков Иван\Knigga.xlsx");
            xlSht = xlWB.ActiveSheet;
            if (dataList.Count == 0)
            {
                MessageBox.Show("Нет данных для экспортан.");
                return;
            }
            for (int i = 0; i < dataList.Count; i++)
            {
                xlSht.Cells[i + 1, 1] = dataList[i];
            }
        }

        private void Button_Formyla_Click(object sender, RoutedEventArgs e)
        {
            string formyla = Text2.Text;
            xlSht.Cells[1, 2].Value = "Формула: ";
            xlSht.Cells[1, 3].FormulaLocal = formyla;
        }
    }
}