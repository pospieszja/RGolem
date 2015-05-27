using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RGolemAddin.Config
{
    static class DataBaseConnection
    {
        static string connectionString =
                                "User=SYSDBA;" +
                                "Password=masterkey;" +
                                "Database=golem_data_develop;" +
                                "DataSource=192.168.5.70;" +
                                "Port=3050";

        public static string GetConnectionString()
        {
            return connectionString;
        }
    }
}
