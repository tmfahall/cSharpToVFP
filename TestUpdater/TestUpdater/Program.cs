using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestUpdater
{
    class Program
    {
        static void updatePhotoNo(string orderNumber, string orderType, decimal photoNo)
        {
            OleDbConnection connection = new OleDbConnection(
            //######## PRODUCTION / TEST #######
            //"Provider=VFPOLEDB.1;Data Source=\\\\ptmserverprime\\E\\vision");
            "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop\\vision");


            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();

                string Update = String.Format("UPDATE WO SET PHOTONO = {0} WHERE ORDERTYPE == '{1}' AND WO_NUMBER == '{2}'", photoNo.ToString(), orderType, orderNumber);

                OleDbCommand cmd3 = new OleDbCommand(Update, connection);

                cmd3.ExecuteNonQuery();

                connection.Close();
            }
        }

        static void Main(string[] args)
        {
            decimal newNum = 1;
            updatePhotoNo("1710", "E", newNum);
        }
    }
}
