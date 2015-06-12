using FirebirdSql.Data.FirebirdClient;
using RGolemAddin.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace RGolemAddin.View
{
    public partial class Form3 : Form
    {
        public DateTime DateTimeFrom { get; set; }
        public DateTime DateTimeTo { get; set; }

        public Form3()
        {
            InitializeComponent();
            SetDateTimePickers();
        }

        private void SetDateTimePickers()
        {
            var yesterday = DateTime.Now.AddDays(-1);
            dateTimePicker.Value = new DateTime(yesterday.Year, yesterday.Month, yesterday.Day);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            long newRow;

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            var xlRange = (Excel.Range)activeWorksheet.Cells[activeWorksheet.Rows.Count, 1];
            long lastRow = xlRange.get_End(Excel.XlDirection.xlUp).Row;

            DateTimeFrom = new DateTime(dateTimePicker.Value.Year, dateTimePicker.Value.Month, dateTimePicker.Value.Day, 6, 0, 0);
            DateTimeTo = new DateTime(dateTimePicker.Value.AddDays(1).Year, dateTimePicker.Value.AddDays(1).Month, dateTimePicker.Value.AddDays(1).Day, 6, 0, 0);

            for (int i = 1; i <= lastRow; i++)
            {
                var cellValue = activeWorksheet.Cells[i, 1].Value2;
                var numberFormat = activeWorksheet.Cells[i, 1].NumberFormat;

                if (cellValue != null && numberFormat == "rrrr-mm-dd")
                {
                    var date = DateTime.FromOADate(cellValue);
                    if (date.Date == DateTimeFrom.Date)
                    {
                        MessageBox.Show("Dzień " + DateTimeFrom.Date.ToShortDateString() + " znajduje się już na liście");
                        return;
                    }
                }
            }

            if (lastRow == 3)
            {
                newRow = 4;
            }
            else
            {
                newRow = lastRow + 9;
                activeWorksheet.get_Range("A4:AE11").Select();
                Globals.ThisAddIn.Application.Selection.Copy(activeWorksheet.Cells[newRow, 1]);
            }

            activeWorksheet.Cells[newRow, 1].Value2 = DateTimeFrom.ToShortDateString();
            activeWorksheet.Cells[newRow, 2].Value2 = WeekOfYearISO8601(DateTimeFrom);

            getDataFromDB(newRow);

            this.Close();
        }

        private static string WeekOfYearISO8601(DateTime date)
        {
            var day = (int)CultureInfo.CurrentCulture.Calendar.GetDayOfWeek(date);
            return date.Year.ToString() + CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(date.AddDays(4 - (day == 0 ? 7 : day)), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString("D2");
        }

        private void getDataFromDB(long newRow)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            //Parametry zapytania SQL
            var paramDateFrom = new FbParameter();
            var paramDateTo = new FbParameter();

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                paramDateFrom.ParameterName = "@czasOd";
                paramDateFrom.Value = DateTimeFrom;
                command.Parameters.Add(paramDateFrom);

                paramDateTo.ParameterName = "@czasDo";
                paramDateTo.Value = DateTimeTo;
                command.Parameters.Add(paramDateTo);

                command.CommandText = @"select sum(d_time) as sum_d_time
                                             , sum(d_tpp) + sum(d_tp) + sum(d_tu) as sum_d_tpp 
                                             , sum(d_tpnp) as sum_d_tpnp
                                             , sum(d_ta) as sum_d_ta
                                             , sum(d_tmp) as sum_d_tmp
                                             , sum(d_g) as sum_d_g
                                             , sum(d_brak) as sum_d_brak
                                             , sum(d_time) + sum(d_tnone) + sum(d_tpp) + sum(d_tpnp) + sum(d_tp) + sum(d_tu) + sum(d_ta) + sum(d_tmp) as sum_d_total
                                             , z
                                             , sv
                                        from raporth 
                                        where czas >= @czasOd and czas < @czasDo and sv in (4,1,2,5,6,7,8)
                                        group by sv, z
                                        order by sv, z";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);
                connection.Close();
            }

            byte svNumber, shiftNumber;
            int totalTime, workTime, tppTime, tpnpTime, taTime, tmpTime, quantity, badQuantity;
            byte columnOffset = 0;

            foreach (DataRow row in dt.Rows)
            {
                svNumber = Convert.ToByte(row["SV"]);
                shiftNumber = Convert.ToByte(row["Z"]);
                totalTime = Convert.ToInt32(row["SUM_D_TOTAL"]);
                workTime = Convert.ToInt32(row["SUM_D_TIME"]);
                tppTime = Convert.ToInt32(row["SUM_D_TPP"]);
                tpnpTime = Convert.ToInt32(row["SUM_D_TPNP"]);
                taTime = Convert.ToInt32(row["SUM_D_TA"]);
                tmpTime = Convert.ToInt32(row["SUM_D_TMP"]);
                quantity = Convert.ToInt32(row["SUM_D_G"]);
                badQuantity = Convert.ToInt32(row["SUM_D_BRAK"]);

                switch (svNumber)
                {
                    case 4:
                        columnOffset = 0;
                        break;
                    case 1:
                        columnOffset = 4;
                        break;
                    case 2:
                        columnOffset = 8;
                        break;
                    case 5:
                        columnOffset = 12;
                        break;
                    case 6:
                        columnOffset = 16;
                        break;
                    case 7:
                        columnOffset = 20;
                        break;
                    case 8:
                        columnOffset = 24;
                        break;
                    default:
                        break;
                }

                if (shiftNumber == 1)
                {
                    // praca %
                    ((Excel.Range)activeWorksheet.Cells[newRow, 5 + columnOffset]).Value2 = workTime / (double)totalTime;
                    // post.plan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 5 + columnOffset]).Value2 = tppTime / (double)totalTime;
                    // post.plan przekr. %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 5 + columnOffset]).Value2 = (double)getExtendedStatusTime(svNumber, shiftNumber) / (double)totalTime;
                    // post.nieplan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 5 + columnOffset]).Value2 = tpnpTime / (double)totalTime;
                    // awarie %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 5 + columnOffset]).Value2 = taTime / (double)totalTime;
                    // mikroprzestoje %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 5 + columnOffset]).Value2 = tmpTime / (double)totalTime;
                    // ilość
                    ((Excel.Range)activeWorksheet.Cells[newRow + 6, 5 + columnOffset]).Value2 = quantity;
                    // braki
                    ((Excel.Range)activeWorksheet.Cells[newRow + 7, 5 + columnOffset]).Value2 = badQuantity;

                    ((Excel.Range)activeWorksheet.Cells[newRow, 5 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 5 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 5 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 5 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 5 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 5 + columnOffset]).NumberFormat = "0.00%";
                }
                if (shiftNumber == 2)
                {
                    // praca %
                    ((Excel.Range)activeWorksheet.Cells[newRow, 6 + columnOffset]).Value2 = workTime / (double)totalTime;
                    // post.plan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 6 + columnOffset]).Value2 = tppTime / (double)totalTime;
                    // post.plan przekr. %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 6 + columnOffset]).Value2 = (double)getExtendedStatusTime(svNumber, shiftNumber) / (double)totalTime;
                    // post.nieplan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 6 + columnOffset]).Value2 = tpnpTime / (double)totalTime;
                    // awarie %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 6 + columnOffset]).Value2 = taTime / (double)totalTime;
                    // mikroprzestoje %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 6 + columnOffset]).Value2 = tmpTime / (double)totalTime;
                    // ilość
                    ((Excel.Range)activeWorksheet.Cells[newRow + 6, 6 + columnOffset]).Value2 = quantity;
                    // braki
                    ((Excel.Range)activeWorksheet.Cells[newRow + 7, 6 + columnOffset]).Value2 = badQuantity;

                    ((Excel.Range)activeWorksheet.Cells[newRow, 6 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 6 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 6 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 6 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 6 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 6 + columnOffset]).NumberFormat = "0.00%";
                }
                if (shiftNumber == 3)
                {
                    // praca %
                    ((Excel.Range)activeWorksheet.Cells[newRow, 7 + columnOffset]).Value2 = workTime / (double)totalTime;
                    // post.plan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 7 + columnOffset]).Value2 = tppTime / (double)totalTime;
                    // post.plan przekr. %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 7 + columnOffset]).Value2 = (double)getExtendedStatusTime(svNumber, shiftNumber) / (double)totalTime;
                    // post.nieplan %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 7 + columnOffset]).Value2 = tpnpTime / (double)totalTime;
                    // awarie %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 7 + columnOffset]).Value2 = taTime / (double)totalTime;
                    // mikroprzestoje %
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 7 + columnOffset]).Value2 = tmpTime / (double)totalTime;
                    // ilość
                    ((Excel.Range)activeWorksheet.Cells[newRow + 6, 7 + columnOffset]).Value2 = quantity;
                    // braki
                    ((Excel.Range)activeWorksheet.Cells[newRow + 7, 7 + columnOffset]).Value2 = badQuantity;

                    ((Excel.Range)activeWorksheet.Cells[newRow, 7 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 1, 7 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 2, 7 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 3, 7 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 4, 7 + columnOffset]).NumberFormat = "0.00%";
                    ((Excel.Range)activeWorksheet.Cells[newRow + 5, 7 + columnOffset]).NumberFormat = "0.00%";
                }
            }
        }

        public class TPZData
        {
            public int st_r_no { get; set; }
            public int st_r_t { get; set; }
            public int d_sr { get; set; }
            public int c_sr { get; set; }
        }

        private int getExtendedStatusTime(byte svNumber, byte shiftNumber)
        {
            // ST_NO - definicja
            // 1 - praca
            // 2 - mikro przestój
            // 3 - nieoznaczony
            // 4 - postój planowany
            // 5 - postój nieplanowany
            // 6 - przezbrajanie
            // 7 - ustawianie
            // 8 - awaria
            //
            // Interesują nas postoje planowane, czyli 4, 6, 7
            // Pobieramy czasy tylko gdy postój jest dłuższy niż wzorcowy

            var extendedStatusTime = 0;

            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            var tpzDataList = new List<TPZData>();

            //Parametry zapytania SQL
            var paramDateFrom = new FbParameter();
            var paramDateTo = new FbParameter();
            var paramSVNumber = new FbParameter();
            var paramShiftNumber = new FbParameter();

            paramDateFrom.ParameterName = "@czasOd";
            paramDateFrom.Value = DateTimeFrom;
            command.Parameters.Add(paramDateFrom);

            paramDateTo.ParameterName = "@czasDo";
            paramDateTo.Value = DateTimeTo;
            command.Parameters.Add(paramDateTo);

            paramShiftNumber.ParameterName = "@z";
            paramShiftNumber.Value = shiftNumber;
            command.Parameters.Add(paramShiftNumber);

            paramSVNumber.ParameterName = "@sv";
            paramSVNumber.Value = svNumber;
            command.Parameters.Add(paramSVNumber);

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select st_r_no
                                             , st_r_t
                                        from tpz
                                        where sv = @sv and st_no in (4,6,7)
                                        order by st_r_no";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);
                connection.Close();
            }

            foreach (DataRow row in dt.Rows)
            {
                var item = new TPZData();
                item.st_r_no = Convert.ToInt32(row["ST_R_NO"]);
                item.st_r_t = Convert.ToInt32(row["ST_R_T"]);
                tpzDataList.Add(item);
            }

            for (var index = 0; index < tpzDataList.Count; index++)
            {
                dt.Clear();
                using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
                {
                    connection.Open();

                    transaction = connection.BeginTransaction();
                    command.Transaction = transaction;
                    command.Connection = connection;

                    command.CommandText = @"select sum(d_sr" + tpzDataList[index].st_r_no + ") as d_sr, sum(c_sr" + tpzDataList[index].st_r_no + ") as c_sr from raporth where sv = @sv and z = @z and czas >= @czasOd and czas < @czasDo";

                    FbDataAdapter adapter = new FbDataAdapter(command);
                    adapter.Fill(dt);

                    connection.Close();
                }

                tpzDataList[index].d_sr = Convert.ToInt32(dt.Rows[0]["D_SR"]);
                tpzDataList[index].c_sr = Convert.ToInt32(dt.Rows[0]["C_SR"]);
            }

            foreach (var item in tpzDataList)
            {
                var refTime = (item.st_r_t * 60) * item.c_sr; // zamiana wzrocowego czasu z minut na sekundy
                var realTime = item.d_sr;

                if (realTime > refTime)
                {
                    extendedStatusTime += realTime - refTime;
                }
            }
            return extendedStatusTime;
        }
    }
}
