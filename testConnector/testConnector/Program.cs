using System;
using System.Data;
using System.Data.OleDb;

/// <summary>
/// There's not a lot of current examples of this that I could find so I'll try to keep this as generic as possible
/// and over comment to help people. 
/// 
/// DBF files can be opened with Open office if they're small enough. I assume MS Office too but I haven't tried.
/// Grab some data from the table once you've opened it to build your SQL query. The table I am targeting is named 
/// WO.DBF and I grabbed an order number and an order type just to test.
/// </summary>

namespace testConnector
{
    class Program
    {
        static DataTable getID(string orderNumber, string orderType)
        {
            //this needs to point to the folder where the db is located, not at the specific file
            OleDbConnection connection = new OleDbConnection(
            "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop\\vision");

            DataTable Result = new DataTable();

            //it may be possible to issue more than one command per open connection, it was unneccessary for me so i haven't tested it.
            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();

                //Here is where to use the info gathered from office
                //select fieldYouWant from dbName where fieldYouCanNarrowItDownFrom == 'specifiData'
                //if the fieldYouCanNarrowItDownFrom is a number it could be saved as a string or any of the number data types. If its a string you need to use it as a string in your query.

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
            //I got these from looking at the db in open office
            string orderNumber = "1710";
            string orderType = "E";


            DataTable data = getID(orderNumber, orderType);

            //this assumes that the data will be in the first row and column and that it can become a string
            string PhotoNo = data.Rows[0][0].ToString();
            Console.WriteLine("PhotoNo = " + PhotoNo);

            //this assumes that the data will be in the first row and column and gets the type
            var test = data.Rows[0][0].GetType();
            Console.WriteLine("Type = " + test);

            //this assumes that the data will be in the first row and column and is an integer and that we want to keep it an integer
            //Console.WriteLine(data.Rows[0].Field<int>(0));

            //this writes out the retrieved datatable to see if the SQL query wasn't specific enough
            foreach (DataRow row in data.Rows)
            {
               foreach(var item in row.ItemArray)
                {
                    Console.WriteLine(item);
                }

            }

            Console.ReadLine();
        }
    }
}
