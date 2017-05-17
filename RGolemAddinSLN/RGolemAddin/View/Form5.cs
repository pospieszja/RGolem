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
    public partial class Form5 : Form
    {
        int Row;

        public Form5()
        {
            InitializeComponent();

            SetDateTimePickers();
        }

        private void SetDateTimePickers()
        {
            var yesterday = DateTime.Now.AddDays(-1);
            dateTimeFrom.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 0, 0, 0);
            dateHourFrom.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 6, 0, 0);
            dateHourTo.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day, 6, 0, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //zerowanie licznika wierszy
            Row = 2;

            Excel.Sheets worksheets = (Excel.Sheets)(Globals.ThisAddIn.Application.Worksheets);

            bool founded = false;
            foreach (Excel.Worksheet item in worksheets)
            {
                if (item.Name == "operatorzy")
                {
                    item.Activate();
                    founded = true;
                }
            }

            if (!founded)
            {
                MessageBox.Show("Nie znaleziono arkusza o nazwie: " + "operatorzy");
                return;
            }

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            activeWorksheet.get_Range("A:B").Clear();
            activeWorksheet.get_Range("A:B").ClearContents();
            generateResult();
        }

        private void generateResult()
        {
            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            var paramDateFrom = new FbParameter();
            var paramDateTo = new FbParameter();

            var dateFrom = new DateTime(dateTimeFrom.Value.Year, dateTimeFrom.Value.Month, dateTimeFrom.Value.Day, dateHourFrom.Value.Hour, 0, 0);
            var dateTo = new DateTime(dateTimeTo.Value.Year, dateTimeTo.Value.Month, dateTimeTo.Value.Day, dateHourTo.Value.Hour, 0, 0);

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            ((Excel.Range)activeWorksheet.Cells[1, 1]).Value2 = "Operator";
            ((Excel.Range)activeWorksheet.Cells[1, 2]).Value2 = "Stanowisko";
            ((Excel.Range)activeWorksheet.Cells[1, 3]).Value2 = "Czas [min]";
            ((Excel.Range)activeWorksheet.Cells[1, 4]).Value2 = "Ilość";
            ((Excel.Range)activeWorksheet.Cells[1, 5]).Value2 = "Braki";
            ((Excel.Range)activeWorksheet.Cells[1, 6]).Value2 = "TPP [min]";
            ((Excel.Range)activeWorksheet.Cells[1, 7]).Value2 = "TPU [min]";

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                paramDateFrom.ParameterName = "@czasOd";
                paramDateFrom.Value = dateFrom;
                command.Parameters.Add(paramDateFrom);

                paramDateTo.ParameterName = "@czasDo";
                paramDateTo.Value = dateTo;
                command.Parameters.Add(paramDateTo);

                command.CommandText = @"select 
                                            uname
                                            , sv
                                            , sum(d_time) as sum_d_time
                                            , sum(d_g) as sum_d_g
                                            , sum(d_brak) as sum_d_brak
                                            , sum(d_tpp) as sum_d_tpp
                                            , sum(d_tpu) as sum_d_tpu
                                        from
                                        (                                        
                                            select u.uname
                                                 , r.sv
                                                 , (r.d_time + r.d_tnone + r.d_tpp + r.d_tpnp + r.d_tp + r.d_tu + r.d_ta + r.d_tmp) as d_time
                                                 , r.d_tpp
                                                 , r.d_tu + r.d_tp as d_tpu
                                                 , r.d_g
                                                 , r.d_brak
                                            from raporth r left outer join user_list u on r.ido = u.ido
                                            where r.czas >= @czasOd and r.czas < @czasDo
                                                and r.ido > 0
                                        ) t
                                        group by uname, sv
                                        order by uname, sv";

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

                    ((Excel.Range)activeWorksheet.Cells[Row, 1]).Value2 = row["UNAME"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 2]).Value2 = svName;
                    ((Excel.Range)activeWorksheet.Cells[Row, 3]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TIME"]) / 60, 2);
                    ((Excel.Range)activeWorksheet.Cells[Row, 4]).Value2 = row["SUM_D_G"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 5]).Value2 = row["SUM_D_BRAK"];
                    ((Excel.Range)activeWorksheet.Cells[Row, 6]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TPP"]) / 60, 2);
                    ((Excel.Range)activeWorksheet.Cells[Row, 7]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TPU"]) / 60, 2);

                    Row += 1;
                }

                connection.Close();
            }

            for (int columnIndex = 1; columnIndex <= 7; columnIndex++)
            {
                ((Excel.Range)activeWorksheet.Cells[1, columnIndex]).EntireColumn.AutoFit();
            }

            this.Close();
        }
    }
}
