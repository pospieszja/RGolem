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
            PopulateListMachine();
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
                command.CommandText = @"select sv, maszyna from maszyny";
                FbDataAdapter adapter = new FbDataAdapter(command);
                adapter.Fill(dt);

                connection.Close();
            }

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

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);


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

            this.Close();
        }
    }
}
