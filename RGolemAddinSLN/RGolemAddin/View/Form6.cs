using FirebirdSql.Data.FirebirdClient;
using RGolemAddin.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace RGolemAddin.View
{
    public partial class Form6 : Form
    {
        public DataTable OrdersDT { get; set; }
        public List<GolemOrder> GolemOrders { get; set; }

        public Task Initialization { get; private set; }

        public Form6()
        {
            Initialization = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            InitializeComponent();

            OrdersDT = await GetDataFromDatabase();
            GolemOrders = new List<GolemOrder>();

            foreach (DataRow order in OrdersDT.Rows)
            {
                GolemOrders.Add(new GolemOrder(order));
            }

            await ExportDataToExcel();
            this.Close();
        }

        private async Task ExportDataToExcel()
        {
            await Task.Delay(100);
            var row = 2;

            Excel.Sheets worksheets = (Excel.Sheets)(Globals.ThisAddIn.Application.Worksheets);

            bool founded = false;
            foreach (Excel.Worksheet item in worksheets)
            {
                if (item.Name == "kolejka zleceń")
                {
                    item.Activate();
                    founded = true;
                }
            }

            if (!founded)
            {
                MessageBox.Show("Nie znaleziono arkusza o nazwie: " + "kolejka zleceń");
                return;
            }

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            activeWorksheet.get_Range("A:C").Clear();
            activeWorksheet.get_Range("A:C").ClearContents();

            ((Excel.Range)activeWorksheet.Cells[1, 1]).Value2 = "Nr zlecenia";
            ((Excel.Range)activeWorksheet.Cells[1, 2]).Value2 = "OCC";
            ((Excel.Range)activeWorksheet.Cells[1, 3]).Value2 = "Stanowisko";

            foreach (var item in GolemOrders)
            {
                ((Excel.Range)activeWorksheet.Cells[row, 1]).Value2 = item.OrderNumber;
                ((Excel.Range)activeWorksheet.Cells[row, 2]).Value2 = item.OCC;
                ((Excel.Range)activeWorksheet.Cells[row, 3]).Value2 = item.SVName;
                row += 1;
            }

            for (int columnIndex = 1; columnIndex <= 3; columnIndex++)
            {
                ((Excel.Range)activeWorksheet.Cells[1, columnIndex]).EntireColumn.AutoFit();
            }
        }

        private async Task<DataTable> GetDataFromDatabase()
        {
            await Task.Delay(100);
            var dt = new DataTable();
            FbTransaction transaction;
            FbCommand command = new FbCommand();

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select 
                                            sv
                                            , zlecenie
                                            , occ
                                        from kolejkaz";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);
                connection.Close();
            }

            return dt;
        }
    }

    public class GolemOrder
    {
        public string OrderNumber { get; set; }
        public int OCC { get; set; }
        public string SVName { get; set; }

        public GolemOrder(DataRow row)
        {
            OrderNumber = SetOrderNumber(Convert.ToString(row["ZLECENIE"]));
            OCC = Convert.ToInt32(row["OCC"]);
            SVName = SetSVName(Convert.ToInt32(row["SV"]));
        }

        private string SetOrderNumber(string orderNumber)
        {
            return orderNumber.Split('/')[0];
        }

        public string SetSVName(int sv)
        {
            string svName = String.Empty;

            switch (sv)
            {
                case 3:
                    svName = "Rozwijarka";
                    break;
                case 4:
                    svName = "Cięcie";
                    break;
                case 1:
                    svName = "Prasa duża";
                    break;
                case 2:
                    svName = "Prasa mała";
                    break;
                case 5:
                    svName = "Spawarka";
                    break;
                case 6:
                    svName = "Robot 1 Linia 1";
                    break;
                case 7:
                    svName = "Robot 1 Linia 2";
                    break;
                case 8:
                    svName = "Robot 2 Linia 4";
                    break;
                case 9:
                    svName = "Szlifierka";
                    break;
                default:
                    break;
            }
            return svName;
        }
    }
}