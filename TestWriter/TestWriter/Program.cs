using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestWriter
{
    class Program
    {
        //

        static void write(string NAME, string ACCOUNT, string ID, DateTime DATE, string TIME, string DOCTYPE, string REFERENCE, string NOTE, string DATA, string FAXDATA, string BIN, string MLTSTORE, string NOTE2, string ORDERTYPE, string FILENAME, string TAG, string REFERENCE2)
        {
            OleDbConnection connection = new OleDbConnection(
            "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop");

            string MyInsert = "insert into DOCUMENT (NAME, ACCOUNT, ID, DATE, TIME, DOCTYPE, REFERENCE, NOTE, DATA, FAXDATA, BIN, MLTSTORE, NOTE2, ORDERTYPE, FILENAME, TAG, REFERENCE2) " +
                "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            OleDbDataAdapter DA = new OleDbDataAdapter();
            
            //
            //

            connection.Open();
            DateTime dt = DateTime.Now;

            if (connection.State == ConnectionState.Open)
            {
                OleDbCommand cmd3 = new OleDbCommand(MyInsert, connection);
                cmd3.Parameters.AddWithValue("NAME", NAME);
                cmd3.Parameters.AddWithValue("ACCOUNT", ACCOUNT);
                cmd3.Parameters.AddWithValue("ID", ID);
                cmd3.Parameters.AddWithValue("DATE", DateTime.Today);
                cmd3.Parameters.AddWithValue("TIME", DATE);
                cmd3.Parameters.AddWithValue("DOCTYPE", DOCTYPE);
                cmd3.Parameters.AddWithValue("REFERENCE", REFERENCE);
                cmd3.Parameters.AddWithValue("NOTE", NOTE);
                cmd3.Parameters.AddWithValue("DATA", DATA);
                cmd3.Parameters.AddWithValue("FAXDATA", FAXDATA);
                cmd3.Parameters.AddWithValue("BIN", BIN);
                cmd3.Parameters.AddWithValue("MLTSTORE", MLTSTORE);
                cmd3.Parameters.AddWithValue("NOTE2", NOTE2);
                cmd3.Parameters.AddWithValue("ORDERTYPE", ORDERTYPE);
                cmd3.Parameters.AddWithValue("FILENAME", FILENAME);
                cmd3.Parameters.AddWithValue("TAG", TAG);
                cmd3.Parameters.AddWithValue("REFERENCE2", REFERENCE2);
                cmd3.ExecuteNonQuery();

                connection.Close();
            }

        }

        static void Main(string[] args)
        {

            string NAME = "test";
            string ACCOUNT = "test";
            string ID = "test";
            DateTime DATE = DateTime.Today;
            string TIME = "test";
            string DOCTYPE = "test";
            string REFERENCE = "test";
            string NOTE = "test";
            string DATA = "test";
            string FAXDATA = "test";
            string BIN = "test";
            string MLTSTORE = "test";
            string NOTE2 = "test";
            string ORDERTYPE = "test";
            string FILENAME = "test";
            string TAG = "test";
            string REFERENCE2 = "test";

            write(NAME, ACCOUNT, ID, DATE, TIME, DOCTYPE, REFERENCE, NOTE, DATA, FAXDATA, BIN, MLTSTORE, NOTE2, ORDERTYPE, FILENAME, TAG, REFERENCE2);
            
            //

            Console.ReadLine();
        }
    }
}
