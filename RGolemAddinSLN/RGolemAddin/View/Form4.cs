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
    public partial class Form4 : Form
    {
        int Row;

        public Form4()
        {
            InitializeComponent();

            SetDateTimePickers();
            PopulateListMachine();
        }

        private void SetDateTimePickers()
        {
            var yesterday = DateTime.Now.AddDays(-1);
            dateTimeFrom.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 0, 0, 0);
            dateHourFrom.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 6, 0, 0);
            dateHourTo.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 6, 0, 0);
        }

        private void PopulateListMachine()
        {
            var dict = new Dictionary<int, string>();

            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;
                command.CommandText = @"select sv, maszyna from maszyny where sv in (1,2,4,5,6,7,8)";
                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            // Dodanie pozycji 0 jako wszystkie
            dict.Add(0, "wszystkie");

            foreach (DataRow row in dt.Rows)
            {
                dict.Add((int)row["SV"], row["MASZYNA"].ToString());
            }

            cbxListMachine.DataSource = new BindingSource(dict, null);
            cbxListMachine.DisplayMember = "Value";
            cbxListMachine.ValueMember = "Key";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int choosenSV;

            //zerowanie licznika wierszy
            Row = 2;

            choosenSV = Convert.ToInt32(cbxListMachine.SelectedValue);
            var machineList = new int[] { 4, 1, 2, 5, 6, 7, 8 };

            // 0 - generuje wynik dla wszystkich maszyn
            // W pozostałych przypadkach generuje wynik dla wybranej maszyny
            if (choosenSV == 0)
            {
                foreach (var svNumber in machineList)
                {
                    generateResultBySV(svNumber);
                }
            }
            else
            {
                generateResultBySV(choosenSV);
            }
        }

        private void generateResultBySV(int svNumber)
        {
            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            var paramMachine = new FbParameter();
            var paramDateFrom = new FbParameter();
            var paramDateTo = new FbParameter();
            var paramShift = new FbParameter();

            var dateFrom = new DateTime(dateTimeFrom.Value.Year, dateTimeFrom.Value.Month, dateTimeFrom.Value.Day, dateHourFrom.Value.Hour, 0, 0);
            var dateTo = new DateTime(dateTimeTo.Value.Year, dateTimeTo.Value.Month, dateTimeTo.Value.Day, dateHourTo.Value.Hour, 0, 0);

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            ((Excel.Range)activeWorksheet.Cells[1, 1]).Value2 = "Nr zlecenia";
            ((Excel.Range)activeWorksheet.Cells[1, 2]).Value2 = "Stanowisko";
            ((Excel.Range)activeWorksheet.Cells[1, 3]).Value2 = "Czas [min]";
            ((Excel.Range)activeWorksheet.Cells[1, 4]).Value2 = "Ilość";
            ((Excel.Range)activeWorksheet.Cells[1, 5]).Value2 = "Braki";
            ((Excel.Range)activeWorksheet.Cells[1, 6]).Value2 = "Zmiana";
            ((Excel.Range)activeWorksheet.Cells[1, 7]).Value2 = "Data";

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                paramMachine.ParameterName = "@sv";
                paramMachine.Value = svNumber;
                command.Parameters.Add(paramMachine);

                paramDateFrom.ParameterName = "@czasOd";
                paramDateFrom.Value = dateFrom;
                command.Parameters.Add(paramDateFrom);

                paramDateTo.ParameterName = "@czasDo";
                paramDateTo.Value = dateTo;
                command.Parameters.Add(paramDateTo);

                command.CommandText = @"select left(s.nazwa,9) as zlecenie
                                             , r.sv as sv
                                             , sum(r.d_time) as sum_d_time
                                             , sum(r.d_g) as sum_d_g
                                             , sum(r.d_brak) as sum_d_brak
                                             , r.z
                                        from raporth r left outer join serie s on r.ids = s.id
                                        where r.czas >= @czasOd and r.czas < @czasDo
                                            and r.sv = @sv and s.id is not null
                                        group by s.nazwa, r.sv, r.z
                                        order by r.z";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                foreach (DataRow row in dt.Rows)
                {
                    string svName = "";
                    switch (Convert.ToInt32(row["SV"]))
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
                            svName = "Robot 1";
                            break;
                        case 7:
                            svName = "Robot 2";
                            break;
                        case 8:
                            svName = "Robot 3";
                            break;
                        case 9:
                            svName = "Szlifierka";
                            break;
                        default:
                            break;
                    }

                    ((Excel.Range)activeWorksheet.Cells[Row, 1]).Value2 = row["ZLECENIE"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 2]).Value2 = svName;
                    ((Excel.Range)activeWorksheet.Cells[Row, 3]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TIME"]) / 60, 2);
                    ((Excel.Range)activeWorksheet.Cells[Row, 4]).Value2 = row["SUM_D_G"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 5]).Value2 = row["SUM_D_BRAK"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 6]).Value2 = row["Z"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 7]).Value2 = dateTimeFrom.Value.Date;
                    ((Excel.Range)activeWorksheet.Cells[Row, 7]).NumberFormat = "m/d/yyyy";

                    Row += 1;
                }

                connection.Close();
            }

            for (int columnIndex = 1; columnIndex <= 20; columnIndex++)
            {
                ((Excel.Range)activeWorksheet.Cells[1, columnIndex]).EntireColumn.AutoFit();
            }

            this.Close();
        }
    }
}
