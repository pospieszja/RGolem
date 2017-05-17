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
    public partial class Form1 : Form
    {
        public Form1()
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
                command.CommandText = @"select sv, maszyna from maszyny where sv in (1,2,4,5,6,7,8,9)";
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

            choosenSV = Convert.ToInt32(cbxListMachine.SelectedValue);
            var machineList = new int[] { 4, 1, 2, 5, 6, 7, 8, 9 };

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
            DataTable dt2 = new DataTable();

            double workedTime = 0;
            var paramMachine = new FbParameter();
            var paramDateFrom = new FbParameter();
            var paramDateTo = new FbParameter();
            var paramShift = new FbParameter();

            var dateFrom = new DateTime(dateTimeFrom.Value.Year, dateTimeFrom.Value.Month, dateTimeFrom.Value.Day, dateHourFrom.Value.Hour, 0, 0);
            var dateTo = new DateTime(dateTimeTo.Value.Year, dateTimeTo.Value.Month, dateTimeTo.Value.Day, dateHourTo.Value.Hour, 0, 0);

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

                command.CommandText = @"select sum(d_time) as sum_d_time
                                             , sum(d_tnone) as sum_d_tnone
                                             , sum(d_tpp) as sum_d_tpp 
                                             , sum(d_tpnp) as sum_d_tpnp
                                             , sum(d_tp) as sum_d_tp
                                             , sum(d_tu) as sum_d_tu
                                             , sum(d_ta) as sum_d_ta
                                             , sum(d_tmp) as sum_d_tmp
                                             , sum(d_g) as sum_d_g
                                             , sum(d_brak) as sum_d_brak
                                             , z
                                        from raporth 
                                        where czas >= @czasOd and czas < @czasDo
                                            and sv = @sv
                                        group by sv, z
                                        order by z";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            string sheetName = "";
            switch (svNumber)
            {
                case 3:
                    sheetName = "Rozwijarka";
                    break;
                case 4:
                    sheetName = "Cięcie";
                    break;
                case 1:
                    sheetName = "Prasa duża";
                    break;
                case 2:
                    sheetName = "Prasa mała";
                    break;
                case 5:
                    sheetName = "Spawarka";
                    break;
                case 6:
                    sheetName = "Robot 1 Linia 1";
                    break;
                case 7:
                    sheetName = "Robot 1 Linia 2";
                    break;
                case 8:
                    sheetName = "Robot 2 Linia 4";
                    break;
                case 9:
                    sheetName = "Szlifierka";
                    break;
                default:
                    break;
            }

            Excel.Sheets worksheets = (Excel.Sheets)(Globals.ThisAddIn.Application.Worksheets);

            bool founded = false;
            foreach (Excel.Worksheet item in worksheets)
            {
                if (item.Name == sheetName)
                {
                    item.Activate();
                    founded = true;
                }
            }

            if (!founded)
            {
                MessageBox.Show("Nie znaleziono arkusza o nazwie: " + sheetName);
                return;
            }

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            activeWorksheet.get_Range("A:U").Clear();
            activeWorksheet.get_Range("A:U").ClearContents();

            /*
             * Projekt layoutu ->
            */

            //Nagłówek
            activeWorksheet.get_Range("A1:Z1").Interior.Color = Color.FromArgb(217, 217, 217);
            activeWorksheet.get_Range("A1", "H1").Merge();
            activeWorksheet.get_Range("A1", "H1").Font.Bold = true;
            activeWorksheet.get_Range("A1", "H1").Font.Size = 14;
            activeWorksheet.get_Range("A1").Value2 = "Raport zbiorczy dla maszyny: " + sheetName + " za okres: " + dateFrom.ToString("yyyy-MM-dd HH:mm") + " - " + dateTo.ToString("yyyy-MM-dd HH:mm");

            //Podsumowanie
            activeWorksheet.get_Range("A2", "U100").Font.Size = 9;

            activeWorksheet.get_Range("C2").Value2 = "Czas ogółem";
            activeWorksheet.get_Range("D2").Value2 = "Praca";
            activeWorksheet.get_Range("E2").Value2 = "Mikropostój";
            activeWorksheet.get_Range("F2").Value2 = "Nieoznaczony";
            activeWorksheet.get_Range("G2").Value2 = "Postój plan.";
            activeWorksheet.get_Range("H2").Value2 = "Postój nieplan.";
            activeWorksheet.get_Range("I2").Value2 = "Przezbrajanie";
            activeWorksheet.get_Range("J2").Value2 = "Ustawianie";
            activeWorksheet.get_Range("K2").Value2 = "Awaria";
            activeWorksheet.get_Range("L2").Value2 = "Ilość";
            activeWorksheet.get_Range("M2").Value2 = "Braki";

            activeWorksheet.get_Range("A3").Value2 = "I";
            activeWorksheet.get_Range("A3").Font.Bold = true;
            activeWorksheet.get_Range("B3").Value2 = "[min]";
            activeWorksheet.get_Range("B4").Value2 = "[%]";

            activeWorksheet.get_Range("A5").Value2 = "II";
            activeWorksheet.get_Range("A5").Font.Bold = true;
            activeWorksheet.get_Range("B5").Value2 = "[min]";
            activeWorksheet.get_Range("B6").Value2 = "[%]";

            activeWorksheet.get_Range("A7").Value2 = "III";
            activeWorksheet.get_Range("A7").Font.Bold = true;
            activeWorksheet.get_Range("B7").Value2 = "[min]";
            activeWorksheet.get_Range("B8").Value2 = "[%]";

            //Postój planowany
            activeWorksheet.get_Range("A9:Z9").Interior.Color = Color.FromArgb(217, 217, 217);
            activeWorksheet.get_Range("A9").Value2 = "Postój planowany";
            activeWorksheet.get_Range("A9").Font.Bold = true;

            activeWorksheet.get_Range("A11").Value2 = "I";
            activeWorksheet.get_Range("A11").Font.Bold = true;
            activeWorksheet.get_Range("B11").Value2 = "[min]";
            activeWorksheet.get_Range("B12").Value2 = "[%]";
            activeWorksheet.get_Range("B13").Value2 = "[delta]";

            activeWorksheet.get_Range("A14").Value2 = "II";
            activeWorksheet.get_Range("A14").Font.Bold = true;
            activeWorksheet.get_Range("B14").Value2 = "[min]";
            activeWorksheet.get_Range("B15").Value2 = "[%]";
            activeWorksheet.get_Range("B16").Value2 = "[delta]";

            activeWorksheet.get_Range("A17").Value2 = "III";
            activeWorksheet.get_Range("A17").Font.Bold = true;
            activeWorksheet.get_Range("B17").Value2 = "[min]";
            activeWorksheet.get_Range("B18").Value2 = "[%]";
            activeWorksheet.get_Range("B19").Value2 = "[delta]";

            //Postój nieplanowany
            activeWorksheet.get_Range("A20:Z20").Interior.Color = Color.FromArgb(217, 217, 217);
            activeWorksheet.get_Range("A20").Value2 = "Postój nieplanowany";
            activeWorksheet.get_Range("A20").Font.Bold = true;

            activeWorksheet.get_Range("A22").Value2 = "I";
            activeWorksheet.get_Range("A22").Font.Bold = true;
            activeWorksheet.get_Range("B22").Value2 = "[min]";
            activeWorksheet.get_Range("B23").Value2 = "[%]";
            activeWorksheet.get_Range("B24").Value2 = "[delta]";

            activeWorksheet.get_Range("A25").Value2 = "II";
            activeWorksheet.get_Range("A25").Font.Bold = true;
            activeWorksheet.get_Range("B25").Value2 = "[min]";
            activeWorksheet.get_Range("B26").Value2 = "[%]";
            activeWorksheet.get_Range("B27").Value2 = "[delta]";

            activeWorksheet.get_Range("A28").Value2 = "III";
            activeWorksheet.get_Range("A28").Font.Bold = true;
            activeWorksheet.get_Range("B28").Value2 = "[min]";
            activeWorksheet.get_Range("B29").Value2 = "[%]";
            activeWorksheet.get_Range("B30").Value2 = "[delta]";

            //Postój przezbrajanie
            activeWorksheet.get_Range("A31:Z31").Interior.Color = Color.FromArgb(217, 217, 217);
            activeWorksheet.get_Range("A31").Value2 = "Przezbrajanie";
            activeWorksheet.get_Range("A31").Font.Bold = true;

            activeWorksheet.get_Range("A33").Value2 = "I";
            activeWorksheet.get_Range("A33").Font.Bold = true;
            activeWorksheet.get_Range("B33").Value2 = "[min]";
            activeWorksheet.get_Range("B34").Value2 = "[%]";
            activeWorksheet.get_Range("B35").Value2 = "[delta]";

            activeWorksheet.get_Range("A36").Value2 = "II";
            activeWorksheet.get_Range("A36").Font.Bold = true;
            activeWorksheet.get_Range("B36").Value2 = "[min]";
            activeWorksheet.get_Range("B37").Value2 = "[%]";
            activeWorksheet.get_Range("B38").Value2 = "[delta]";

            activeWorksheet.get_Range("A39").Value2 = "III";
            activeWorksheet.get_Range("A39").Font.Bold = true;
            activeWorksheet.get_Range("B39").Value2 = "[min]";
            activeWorksheet.get_Range("B40").Value2 = "[%]";
            activeWorksheet.get_Range("B41").Value2 = "[delta]";

            //Ustawianie
            activeWorksheet.get_Range("A42:Z42").Interior.Color = Color.FromArgb(217, 217, 217);
            activeWorksheet.get_Range("A42").Value2 = "Ustawianie";
            activeWorksheet.get_Range("A42").Font.Bold = true;

            activeWorksheet.get_Range("A44").Value2 = "I";
            activeWorksheet.get_Range("A44").Font.Bold = true;
            activeWorksheet.get_Range("B44").Value2 = "[min]";
            activeWorksheet.get_Range("B45").Value2 = "[%]";
            activeWorksheet.get_Range("B46").Value2 = "[delta]";

            activeWorksheet.get_Range("A47").Value2 = "II";
            activeWorksheet.get_Range("A47").Font.Bold = true;
            activeWorksheet.get_Range("B47").Value2 = "[min]";
            activeWorksheet.get_Range("B48").Value2 = "[%]";
            activeWorksheet.get_Range("B49").Value2 = "[delta]";

            activeWorksheet.get_Range("A50").Value2 = "III";
            activeWorksheet.get_Range("A50").Font.Bold = true;
            activeWorksheet.get_Range("B50").Value2 = "[min]";
            activeWorksheet.get_Range("B51").Value2 = "[%]";
            activeWorksheet.get_Range("B52").Value2 = "[delta]";

            //Operatorzy
            activeWorksheet.get_Range("P2").Value2 = "I zmiana";
            activeWorksheet.get_Range("R2").Value2 = "II zmiana";
            activeWorksheet.get_Range("T2").Value2 = "III zmiana";

            /*
             * Projekt layoutu <-
            */


            int rowIndex = 3;
            foreach (DataRow row in dt.Rows)
            {


                //string tempValue = ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, 1]).Value2;
                //if (tempValue != null)
                //{
                //    ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, 1]).Value2 = tempValue.Substring(3, tempValue.Length - 3);
                //}

                //całkowity czas
                workedTime = Convert.ToDouble(row["SUM_D_TIME"]) + Convert.ToDouble(row["SUM_D_TMP"]) + Convert.ToDouble(row["SUM_D_TNONE"])
                                + Convert.ToDouble(row["SUM_D_TPP"]) + Convert.ToDouble(row["SUM_D_TPNP"]) + Convert.ToDouble(row["SUM_D_TP"])
                                + Convert.ToDouble(row["SUM_D_TU"]) + Convert.ToDouble(row["SUM_D_TA"]);

                //zamiana całkowitego czasu z sek na min
                workedTime = workedTime / 60;

                ((Excel.Range)activeWorksheet.Cells[rowIndex, 3]).Value2 = Math.Round(workedTime, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 4]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TIME"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 5]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TMP"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 6]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TNONE"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 7]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TPP"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 8]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TPNP"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 9]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TP"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 10]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TU"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 11]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TA"]) / 60, 2);
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 12]).Value2 = row["SUM_D_G"];
                ((Excel.Range)activeWorksheet.Cells[rowIndex, 13]).Value2 = row["SUM_D_BRAK"];

                for (int columnIndex = 4; columnIndex <= 11; columnIndex++)
                {
                    ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, columnIndex]).Value2 = ((Excel.Range)activeWorksheet.Cells[rowIndex, columnIndex]).Value2 / workedTime;
                    ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, columnIndex]).NumberFormat = "0.00%";
                }

                rowIndex += 2;
            }


            //Postój planowany - ST_NO = 4
            dt.Clear();
            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select st_r_no
                                              , st_r_desc
                                              , st_r_t
                                        from tpz
                                        where st_no = 4
                                            and sv = @sv";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            int j = 3;
            foreach (DataRow row in dt.Rows)
            {
                rowIndex = 10;
                ((Excel.Range)activeWorksheet.Cells[rowIndex, j]).Value2 = row["ST_R_DESC"];
                dt2.Clear();
                using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
                {
                    connection.Open();

                    transaction = connection.BeginTransaction();
                    command.Transaction = transaction;
                    command.Connection = connection;

                    command.CommandText = @"select sum(d_sr" + row["ST_R_NO"] + ") as sum_d_sr, sum(c_sr" + row["ST_R_NO"] + ") as sum_c_sr, z from raporth where sv = @sv and czas >= @czasOd and czas < @czasDo group by z order by z";

                    FbDataAdapter adapter = new FbDataAdapter(command);
                    adapter.Fill(dt2);
                    foreach (DataRow row2 in dt2.Rows)
                    {
                        ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, j]).Value2 = Math.Round(Convert.ToDouble(row2["SUM_D_SR"]) / 60, 2);
                        //if (Convert.ToInt32(row2["SUM_C_SR"]) > 0)
                        //{
                        if (Convert.ToDouble(row["ST_R_T"]) > 0 && Convert.ToDouble(row2["SUM_C_SR"]) > 0)
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) / (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])) - 1);
                            }
                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) - (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])));

                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).NumberFormat = "0.00%";
                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).NumberFormat = "0.00";
                            if (((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 > 0)
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            }
                        //}
                        rowIndex += 3;
                    }

                    connection.Close();
                }
                j++;
            }


            //Postój nieplanowany - ST_NO = 5
            dt.Clear();
            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select st_r_no
                                                         , st_r_desc
                                                         , st_r_t
                                                    from tpz
                                                    where st_no = 5
                                                        and sv = @sv";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            j = 3;
            foreach (DataRow row in dt.Rows)
            {
                rowIndex = 21;
                ((Excel.Range)activeWorksheet.Cells[rowIndex, j]).Value2 = row["ST_R_DESC"];
                dt2.Clear();
                using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
                {
                    connection.Open();

                    transaction = connection.BeginTransaction();
                    command.Transaction = transaction;
                    command.Connection = connection;

                    command.CommandText = @"select sum(d_sr" + row["ST_R_NO"] + ") as sum_d_sr, sum(c_sr" + row["ST_R_NO"] + ") as sum_c_sr from raporth where sv = @sv and czas >= @czasOd and czas < @czasDo group by z order by z";

                    FbDataAdapter adapter = new FbDataAdapter(command);
                    adapter.Fill(dt2);
                    foreach (DataRow row2 in dt2.Rows)
                    {
                        ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, j]).Value2 = Math.Round(Convert.ToDouble(row2["SUM_D_SR"]) / 60, 2);
                        //if (Convert.ToInt32(row2["SUM_C_SR"]) > 0)
                        //{
                            if (Convert.ToDouble(row2["SUM_D_SR"]) > 0)
                            {
                                if (Convert.ToDouble(row["ST_R_T"]) > 0)
                                {
                                    ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) / (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])) - 1);
                                }
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) - (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])));
                            }

                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).NumberFormat = "0.00%";
                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).NumberFormat = "0.00";
                            if (((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 > 0)
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            }
                        //}
                        rowIndex += 3;
                    }

                    connection.Close();
                }
                j++;
            }

            //Przezbrojanie - ST_NO = 6
            dt.Clear();
            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select st_r_no
                                                , st_r_desc
                                                , st_r_t
                                        from tpz
                                        where st_no = 6
                                            and sv = @sv";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            j = 3;
            foreach (DataRow row in dt.Rows)
            {
                rowIndex = 32;
                ((Excel.Range)activeWorksheet.Cells[rowIndex, j]).Value2 = row["ST_R_DESC"];
                dt2.Clear();
                using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
                {
                    connection.Open();

                    transaction = connection.BeginTransaction();
                    command.Transaction = transaction;
                    command.Connection = connection;

                    command.CommandText = @"select sum(d_sr" + row["ST_R_NO"] + ") as sum_d_sr, sum(c_sr" + row["ST_R_NO"] + ") as sum_c_sr from raporth where sv = @sv and czas >= @czasOd and czas < @czasDo group by z order by z";

                    FbDataAdapter adapter = new FbDataAdapter(command);
                    adapter.Fill(dt2);
                    foreach (DataRow row2 in dt2.Rows)
                    {
                        ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, j]).Value2 = Math.Round(Convert.ToDouble(row2["SUM_D_SR"]) / 60, 2);
                        //if (Convert.ToInt32(row2["SUM_C_SR"]) > 0)
                        //{
                            if (Convert.ToDouble(row2["SUM_D_SR"]) > 0)
                            {
                                if (Convert.ToDouble(row["ST_R_T"]) > 0)
                                {
                                    ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) / (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])) - 1);
                                }
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) - (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])));
                            }

                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).NumberFormat = "0.00%";
                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).NumberFormat = "0.00";
                            if (((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 > 0)
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            }
                        //}
                        rowIndex += 3;
                    }
                    connection.Close();
                }
                j++;
            }

            //Ustawianie - ST_NO = 7
            dt.Clear();
            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select st_r_no
                                                         , st_r_desc
                                                         , st_r_t
                                                    from tpz
                                                    where st_no = 7
                                                        and sv = @sv";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }


            j = 3;
            foreach (DataRow row in dt.Rows)
            {
                rowIndex = 43;
                ((Excel.Range)activeWorksheet.Cells[rowIndex, j]).Value2 = row["ST_R_DESC"];
                dt2.Clear();
                using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
                {
                    connection.Open();

                    transaction = connection.BeginTransaction();
                    command.Transaction = transaction;
                    command.Connection = connection;

                    command.CommandText = @"select sum(d_sr" + row["ST_R_NO"] + ") as sum_d_sr, sum(c_sr" + row["ST_R_NO"] + ") as sum_c_sr from raporth where sv = @sv and czas >= @czasOd and czas < @czasDo group by z order by z";

                    FbDataAdapter adapter = new FbDataAdapter(command);
                    adapter.Fill(dt2);
                    foreach (DataRow row2 in dt2.Rows)
                    {
                        ((Excel.Range)activeWorksheet.Cells[rowIndex + 1, j]).Value2 = Math.Round(Convert.ToDouble(row2["SUM_D_SR"]) / 60, 2);
                        //if (Convert.ToInt32(row2["SUM_C_SR"]) > 0)
                        //{
                            if (Convert.ToDouble(row2["SUM_D_SR"]) > 0)
                            {
                                if (Convert.ToDouble(row["ST_R_T"]) > 0)
                                {
                                    ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) / (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])) - 1);
                                }
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Value2 = ((Convert.ToDouble(row2["SUM_D_SR"]) / 60) - (Convert.ToDouble(row["ST_R_T"]) * Convert.ToDouble(row2["SUM_C_SR"])));
                            }

                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).NumberFormat = "0.00%";
                            ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).NumberFormat = "0.00";
                            if (((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Value2 > 0)
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 2, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                ((Excel.Range)activeWorksheet.Cells[rowIndex + 3, j]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            }
                        //}
                        rowIndex += 3;
                    }

                    connection.Close();
                }
                j++;
            }

            //Czas pracy operatora
            dt.Clear();
            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.CommandText = @"select 
                                            left(uname, position(' ',uname) ) as uname
                                            , sv
                                            , sum(d_time) as sum_d_time
                                            , z
                                        from
                                        (                                        
                                            select u.uname
                                                 , r.sv
                                                 , (r.d_time + r.d_tnone + r.d_tpp + r.d_tpnp + r.d_tp + r.d_tu + r.d_ta + r.d_tmp) as d_time
                                                 , r.z
                                            from raporth r left outer join user_list u on r.ido = u.ido
                                            where r.czas >= @czasOd and r.czas < @czasDo
                                                and r.ido > 0
                                                and r.sv = @sv
                                        ) t
                                        group by uname, z, sv
                                        order by z, sum_d_time desc";

                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                int numberOfOperatorsPerFirstShift = 2;
                int numberOfOperatorsPerSecondShift = 2;
                int numberOfOperatorsPerThirdShift = 2;

                foreach (DataRow row in dt.Rows)
                {
                    int columnNo = 16;

                    switch (Convert.ToInt32(row["Z"]))
                    {
                        case 1:
                            columnNo = 16;
                            numberOfOperatorsPerFirstShift += 1;
                            rowIndex = numberOfOperatorsPerFirstShift;
                            break;
                        case 2:
                            columnNo = 18;
                            numberOfOperatorsPerSecondShift += 1;
                            rowIndex = numberOfOperatorsPerSecondShift;
                            break;
                        case 3:
                            columnNo = 20;
                            numberOfOperatorsPerThirdShift += 1;
                            rowIndex = numberOfOperatorsPerThirdShift;
                            break;
                        default:
                            break;
                    }

                    ((Excel.Range)activeWorksheet.Cells[rowIndex, columnNo]).Value2 = row["UNAME"];
                    ((Excel.Range)activeWorksheet.Cells[rowIndex, columnNo+1]).Value2 = Math.Round(Convert.ToDouble(row["SUM_D_TIME"]) / 60, 2);

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
