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
    public partial class Form2 : Form
    {
        string order;
        string materialNo;
        int quantity;
        string materialDesc;

        public Form2()
        {
            InitializeComponent();
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            var activeCell = Globals.ThisAddIn.Application.ActiveCell;
            var rowNo = activeCell.Row;

            order = Convert.ToString(activeCell.Value2);
            materialNo = Convert.ToString(activeWorksheet.Cells[rowNo, 3].Value2);
            quantity = Convert.ToInt32(activeWorksheet.Cells[rowNo, 12].Value2);
            materialDesc = Convert.ToString(activeWorksheet.Cells[rowNo, 14].Value2);

            tbxOrder.Text = order;
            tbxMaterialNo.Text = materialNo;
            tbxMaterialDesc.Text = materialDesc;

            numQuanSV4.Value = quantity;
            numQuanSV1.Value = quantity;
            numQuanSV2.Value = quantity;
            numQuanSV5.Value = quantity;
            numQuanSV6.Value = quantity;
            numQuanSV7.Value = quantity;
            numQuanSV8.Value = quantity;

            //domyślna wartość OCP
            numOcpSV4.Value = 1;
            numOcpSV1.Value = 1;
            numOcpSV2.Value = 1;
            numOcpSV5.Value = 1;
            numOcpSV6.Value = 1;
            numOcpSV7.Value = 1;
            numOcpSV8.Value = 1;

            //domyslna wartość OCU
            numOcuSV4.Value = 15;
            numOcuSV1.Value = 15;
            numOcuSV2.Value = 15;
            numOcuSV5.Value = 5;
            numOcuSV6.Value = 15;
            numOcuSV7.Value = 15;
            numOcuSV8.Value = 15;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            exportOrdersIntoGolemDatabase();
        }

        private void exportOrdersIntoGolemDatabase()
        {
            FbTransaction transaction;
            FbCommand command = new FbCommand();

            int occ = 0, ocp = 0, ocu = 0;
            string orderSV = String.Empty;

            Dictionary<int, string> svDict = new Dictionary<int, string>();

            //Cięcie
            svDict.Add(4, "C");
            //Prasa duża
            svDict.Add(1, "PD");
            //Prasa mała
            svDict.Add(2, "PM");
            //Spawarka
            svDict.Add(5, "SP");
            //Robot Linia 1
            svDict.Add(6, "L1");
            //Robot Linia 2
            svDict.Add(7, "L2");
            //Robot Linia 4
            svDict.Add(8, "L4");

            using (FbConnection connection = new FbConnection(DataBaseConnection.GetConnectionString()))
            {
                connection.Open();

                transaction = connection.BeginTransaction();
                command.Transaction = transaction;
                command.Connection = connection;

                command.Parameters.Add("@sv", FbDbType.Integer);
                command.Parameters.Add("@materialNo", FbDbType.VarChar);
                command.Parameters.Add("@orderSV", FbDbType.VarChar);
                command.Parameters.Add("@quantity", FbDbType.Integer);
                command.Parameters.Add("@occ", FbDbType.Integer);
                command.Parameters.Add("@ocp", FbDbType.Integer);
                command.Parameters.Add("@ocu", FbDbType.Integer);
                command.Parameters.Add("@materialDesc", FbDbType.VarChar);
                command.Parameters.Add("@zorder", FbDbType.Integer);

                foreach (var item in svDict)
                {
                    switch (item.Key)
                    {
                        case 4:
                            quantity = (int)numQuanSV4.Value;
                            occ = (int)numOccSV4.Value;
                            ocp = (int)numOcpSV4.Value;
                            ocu = (int)numOcuSV4.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 1:
                            quantity = (int)numQuanSV1.Value;
                            occ = (int)numOccSV1.Value;
                            ocp = (int)numOcpSV1.Value;
                            ocu = (int)numOcuSV1.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 2:
                            quantity = (int)numQuanSV2.Value;
                            occ = (int)numOccSV2.Value;
                            ocp = (int)numOcpSV2.Value;
                            ocu = (int)numOcuSV2.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 5:
                            quantity = (int)numQuanSV5.Value;
                            occ = (int)numOccSV5.Value;
                            ocp = (int)numOcpSV5.Value;
                            ocu = (int)numOcuSV5.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 6:
                            quantity = (int)numQuanSV6.Value;
                            occ = (int)numOccSV6.Value;
                            ocp = (int)numOcpSV6.Value;
                            ocu = (int)numOcuSV6.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 7:
                            quantity = (int)numQuanSV7.Value;
                            occ = (int)numOccSV7.Value;
                            ocp = (int)numOcpSV7.Value;
                            ocu = (int)numOcuSV7.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        case 8:
                            quantity = (int)numQuanSV8.Value;
                            occ = (int)numOccSV8.Value;
                            ocp = (int)numOcpSV8.Value;
                            ocu = (int)numOcuSV8.Value;
                            orderSV = order + "/" + item.Value;
                            break;
                        default:
                            break;
                    }

                    if (quantity > 0)
                    {
                        command.CommandText = @"INSERT INTO KOLEJKAZ (SV, PRODUKT, ZLECENIE, ILE_Z, OCC, OC_P, OC_U, NAZWAKZ, ZORDER, ZEND, ZHIDE) VALUES (@sv, @materialNo, @orderSV, @quantity, @occ, @ocp, @ocu, @materialDesc, @zorder, 0, 0);";
                        command.Parameters[0].Value = item.Key;
                        command.Parameters[1].Value = materialNo;
                        command.Parameters[2].Value = orderSV;
                        command.Parameters[3].Value = quantity;
                        command.Parameters[4].Value = occ;
                        command.Parameters[5].Value = ocp;
                        command.Parameters[6].Value = ocu;
                        command.Parameters[7].Value = materialDesc;
                        command.Parameters[8].Value = order;
                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
                connection.Close();
                MessageBox.Show("Eksport zakończony poprawnie");
                this.Close();
            }
        }
    }
}
