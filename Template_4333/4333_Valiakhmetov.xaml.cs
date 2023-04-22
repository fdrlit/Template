using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
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

namespace Template_4333
{
    /// <summary>
    /// Interaction logic for _4333_Valiakhmetov.xaml
    /// </summary>
    public partial class _4333_Valiakhmetov : System.Windows.Window
    {
        public _4333_Valiakhmetov()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (2.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (RentalOrdersssEntities ordersEntities = new RentalOrdersssEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    ordersEntities.reOrders.Add(new reOrder()
                    {
                        Code = list[i, 0],
                        OrderCode = list[i, 1],
                        CreateDate = list[i, 2],
                        OrderTime = list[i, 3],
                        ClientCode = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        CloseDate = list[i, 7],
                        RentalTime = list[i, 8],

                    });
                }
                ordersEntities.SaveChanges();
            }

        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<reOrder>> ByData = new Dictionary<string, List<reOrder>>();
            using (RentalOrdersssEntities usersEntities = new RentalOrdersssEntities())
            {

                var GroupedByData = usersEntities.reOrders.ToList().GroupBy(w => w.Status);


                foreach (var group in GroupedByData)
                {
                    ByData[group.Key] = group.ToList();
                }
            }


            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);


            app.Visible = true;

            foreach (var order in ByData)
            {
                string position = order.Key;
                List<reOrder> orders = order.Value;

                Excel.Worksheet worksheet = app.Worksheets.Add();
                if (position != "")
                {
                    worksheet.Name = position;

                    worksheet.Cells[1, 1] = "Id";
                    worksheet.Cells[1, 2] = "Код заказа";
                    worksheet.Cells[1, 3] = "Дата создания";
                    worksheet.Cells[1, 4] = "Код клиента";
                    worksheet.Cells[1, 5] = "Услуги";
                }
                int rowIndex = 2; // Начальная строка для записи данных
                foreach (reOrder orderss in orders)
                {
                    worksheet.Cells[rowIndex, 1] = orderss.Code;
                    worksheet.Cells[rowIndex, 2] = orderss.OrderCode;
                    worksheet.Cells[rowIndex, 3] = orderss.CreateDate;
                    worksheet.Cells[rowIndex, 4] = orderss.ClientCode;
                    worksheet.Cells[rowIndex, 5] = orderss.Services;
                    rowIndex++;
                }
            }
        }
    }
}