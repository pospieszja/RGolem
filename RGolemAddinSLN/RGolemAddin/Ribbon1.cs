using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using System.Data;
using RGolemAddin.Config;
using RGolemAddin.View;

namespace RGolemAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            FbTransaction transaction;
            FbCommand command = new FbCommand();
            DataTable dt = new DataTable();

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;
                command.CommandText = @"select sv, maszyna from maszyny order by sv";
                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

            Excel.Window window = e.Control.Context;
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);


            ((Excel.Range)activeWorksheet.Cells[1, 1]).Value2 = "SV";
            ((Excel.Range)activeWorksheet.Cells[1, 2]).Value2 = "MASZYNA";
            int i = 2;
            foreach (DataRow row in dt.Rows)
            {
                ((Excel.Range)activeWorksheet.Cells[i, 1]).Value2 = row["SV"];
                ((Excel.Range)activeWorksheet.Cells[i, 2]).Value2 = row["MASZYNA"];
                i++;
            }
            ((Excel.Range)activeWorksheet.Cells[1, 2]).EntireColumn.AutoFit();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form1();
            form.ShowDialog();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            string workbookName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            string worksheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (workbookName == "zlecenie_prod_soutec_planGhant.xlsm" && worksheetName == "zlecenia do realizacji" && activeCell.Column == 5)
            {
                var form = new Form2();
                form.Show();
            }
            else
            {
                MessageBox.Show("Akcja niemożliwa do wykonania: nieprawidłowy arkusz lub zeszyt");
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string worksheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
            if (worksheetName == "dane zbiorcze")
            {
                var form = new Form3();
                form.Show();
            }
            else
            {
                MessageBox.Show("Raport można wykonać tylko w arkuszu 'dane zbiorcze'");
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form4();
            form.Show();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form5();
            form.Show();
        }
    }
}