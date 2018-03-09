using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testConnector
{
    class Program
    {
        static DataTable getID(string orderNumber, string orderType)
        {
            OleDbConnection connection = new OleDbConnection(
            "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop\\vision");

            DataTable Result = new DataTable();

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string mySQL = String.Format("SELECT PHOTONO FROM WO WHERE ORDERTYPE == '{0}' AND WO_NUMBER == '{1}'", orderType, orderNumber);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result;
        }

        static void Main(string[] args)
        {
            string orderNumber = "1710";
            string orderType = "E";
            DataTable data = getID(orderNumber, orderType); //-- > get from work order

            string NAME = data.Rows[0][0].ToString();
            var test = data.Rows[0][0].GetType();
            foreach (DataRow row in data.Rows)
            {
               foreach(var item in row.ItemArray)
                {
                    Console.WriteLine(item);
                }

            }

            Console.WriteLine("NAME = " + NAME);
            Console.WriteLine("Type = " + test);
            //Console.WriteLine(data.Rows[0].Field<int>(0));

            Console.ReadLine();
        }
    }
}
