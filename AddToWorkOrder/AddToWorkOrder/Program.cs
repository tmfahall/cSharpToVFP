using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;

namespace AddToWorkOrder
{

    class Program
    {
        static string connectionReadyForProduction()
        {
            bool isReady = true;

            if (isReady == false)
            {
                return "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop\\vision";
            }
            else
            {
                return "Provider=VFPOLEDB.1;Data Source=\\\\ptmserverprime\\E\\vision";
            }
        }

        static string cameraReadyForProduction()
        {
            bool isReady = true;

            if (isReady == false)
            {
                return @"C:\Users\Administrator\Pictures";
            }
            else
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string folderToCheck = String.Format("{0}//ToUpload", desktop);
                return folderToCheck;
            }
        }

        static string imageDestinationReadyForProduction()
        {
            bool isReady = true;

            if (isReady == false)
            {
                return @"C:\Users\Administrator\Desktop\visionTest";
            }
            else
            {
                return @"\\ptmserverprime\E\vision\EDM";
            }
        }

        static int getIdNumber()
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());

            DataTable Result = new DataTable();

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string mySQL = String.Format("SELECT MAX(CAST(ID AS Int)) FROM DOCUMENT");

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }
            return Result.Rows[0].Field<int>(0);

        }

        //NEED TO DO MORE RESEARCH BEFORE REINDEXING, ERRORS OUT NOW

        //static void indexDb()
        //{
        //    OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());
        //    DataTable Result = new DataTable();
        //    connection.Open();

        //    if (connection.State == ConnectionState.Open)
        //    {
        //        try
        //        {
        //            OleDbDataAdapter DA = new OleDbDataAdapter();
        //            string mySQL = "INDEX ON ACCOUNT TO DOCACCT.IDX ASCENDING";

        //            OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

        //            MyQuery.ExecuteNonQuery();

        //            connection.Close();
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("Error creating DOCACCT.IDX");
        //            Console.WriteLine(e.ToString());
        //        }

        //        try
        //        {
        //            OleDbDataAdapter DA = new OleDbDataAdapter();
        //            string mySQL = "INDEX ON DATE TO DOCDATE.IDX ASCENDING";

        //            OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

        //            MyQuery.ExecuteNonQuery();

        //            connection.Close();
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("Error creating DOCDATE.IDX");
        //            Console.WriteLine(e.ToString());
        //        }

        //        try
        //        {
        //            OleDbDataAdapter DA = new OleDbDataAdapter();
        //            string mySQL = "INDEX ON ID TO DOCID.IDX ASCENDING";

        //            OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

        //            MyQuery.ExecuteNonQuery();

        //            connection.Close();
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("Error creating DOCID.IDX");
        //            Console.WriteLine(e.ToString());
        //        }
        //    }
        //}

        static string getAccount(string orderNumber, string orderType)
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());

            DataTable Result = new DataTable();

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string mySQL = String.Format("SELECT ACCOUNT FROM WO WHERE WO_NUMBER == '{0}' AND ORDERTYPE == '{1}'", orderNumber, orderType);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result.Rows[0][0].ToString();
        }

        static string getName(string orderNumber, string orderType)
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());


            DataTable Result = new DataTable();

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string mySQL = String.Format("SELECT NAME FROM WO WHERE WO_NUMBER == '{0}' AND ORDERTYPE == '{1}'", orderNumber, orderType);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result.Rows[0][0].ToString();
        }

        static void welcomeMessage()
        {
            Console.WriteLine("This application will transfer images to a work order.");
        }

        static void checkForImages()
        {
            Console.WriteLine("Please ensure the camera is connected and turned on.");
            Console.WriteLine("Then drag the pictures to the folder ToUpload on the desktop.");
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
            Console.Clear();

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderToCheck = String.Format("{0}//ToUpload", desktop);
            
            var timeUntilNextTry = TimeSpan.FromSeconds(3);
            var timeUntilGiveUp = DateTime.Now.Add(TimeSpan.FromMinutes(2));
            Console.Write("Searching for images");

            bool isCamera()
            {
                if (File.Exists(folderToCheck))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            if (isCamera() == false)
            {
                Console.Clear();

                if (DateTime.Now > timeUntilGiveUp)
                {
                    Console.WriteLine("Images could not be found. Please make sure images were added to the correct folder");
                    Console.WriteLine("Press any key to exit");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                Thread.Sleep(timeUntilNextTry);
                isCamera();
            }
            else
            {
                Console.WriteLine("\r\nImages Detected");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
                Console.Clear();
            }
        }

        static string getOrderType()
        {
            bool finished = false;
            string orderType = "";

            while (!finished)
            {
                Console.WriteLine("Please select type");
                Console.WriteLine("1. Work Order");
                Console.WriteLine("2. Inhouse");
                Console.WriteLine("3. Estimate");
                Console.WriteLine("Q. Quit");
                string selectionChoice = Console.ReadLine();
                switch (selectionChoice)
                {
                    case "1":
                        orderType = "W";
                        finished = true;
                        break;
                    case "2":
                        orderType = "I";
                        finished = true;
                        break;
                    case "3":
                        orderType = "E";
                        finished = true;
                        break;
                    case "q":
                        finished = true;
                        Environment.Exit(0);
                        break;

                        //check edge cases!!
                }
            }

            return orderType;
        }

        static string buildReference(int previousPicsCount, int iteratorForOrders, string orderNumber)
        {
            int numberToUse = previousPicsCount + iteratorForOrders;

            string reference = String.Format("{0}-{1}", orderNumber, numberToUse.ToString());

            return reference;
        }

        static int getPhotoNo(string orderNumber, string orderType)
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());

            DataTable Result = new DataTable();

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string mySQL = String.Format("SELECT PHOTONO FROM WO WHERE WO_NUMBER == '{0}' AND ORDERTYPE == '{1}'", orderNumber, orderType);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Convert.ToInt16(Result.Rows[0][0]);
        }

        static void updatePhotoNo(string orderNumber, string orderType, decimal photoNo)
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());

            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();
                string Update = String.Format("UPDATE WO SET PHOTONO = {0} WHERE ORDERTYPE == '{1}' AND WO_NUMBER == '{2}'", photoNo, orderType, orderNumber);

                OleDbCommand MyQuery = new OleDbCommand(Update, connection);

                MyQuery.ExecuteNonQuery();

                connection.Close();
            }
        }

        static string getOrderNumber(string orderType)
        {
            string orderTypeFull = "";
            bool orderNumberCorrect = false;
            string orderNumber = "";

            switch (orderType)
            {
                case "W":
                    orderTypeFull = "work order";
                    break;
                case "I":
                    orderTypeFull = "inhouse order";
                    break;
                case "E":
                    orderTypeFull = "estimate";
                    break;
            }

            while (!orderNumberCorrect)
            {
                Console.WriteLine("What is the {0} number", orderTypeFull);

                var inputOrderNumber = Console.ReadLine();

                Console.WriteLine("You selected {0} number {1}, is this correct? y/n", orderTypeFull, inputOrderNumber);
                string answer = Console.ReadLine();
                if (answer.ToLower() == "y")
                {
                    orderNumber = inputOrderNumber;
                    orderNumberCorrect = true;
                }
                else
                {
                    getOrderNumber(orderType);
                }
            }

            return orderNumber;
        }

        static string getRandom(Random random)
        {
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var stringChars = new char[3];

            for (int i = 0; i < stringChars.Length; i++)
            {
                stringChars[i] = chars[random.Next(chars.Length)];
            }

            return new string(stringChars);
        }

        static void appendDb(string NAME, string ACCOUNT, string ID, DateTime DATE, string TIME, string DOCTYPE, string REFERENCE, string NOTE, string DATA, string FAXDATA, string BIN, string MLTSTORE, string NOTE2, string ORDERTYPE, string FILENAME, string TAG, string REFERENCE2)
        {
            OleDbConnection connection = new OleDbConnection(connectionReadyForProduction());

            string MyInsert = "insert into DOCUMENT (NAME, ACCOUNT, ID, DATE, TIME, DOCTYPE, REFERENCE, NOTE, DATA, FAXDATA, BIN, MLTSTORE, NOTE2, ORDERTYPE, FILENAME, TAG, REFERENCE2) " +
                "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            OleDbDataAdapter DA = new OleDbDataAdapter();

            connection.Open();


            if (connection.State == ConnectionState.Open)
            {
                OleDbCommand cmd3 = new OleDbCommand(MyInsert, connection);
                cmd3.Parameters.AddWithValue("NAME", NAME);
                cmd3.Parameters.AddWithValue("ACCOUNT", ACCOUNT);
                cmd3.Parameters.AddWithValue("ID", ID);
                cmd3.Parameters.AddWithValue("DATE", DATE);
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



        //####### BODY
        static void Main(string[] args)
        {
            welcomeMessage();
            checkForImages();

            string picturesOnCamera = cameraReadyForProduction();
            Console.WriteLine("Moving pictures from = " + picturesOnCamera);

            string imageDestination = imageDestinationReadyForProduction();
            Console.WriteLine("Moving files to = " + imageDestination);

            string[] pictures = Directory.GetFiles(picturesOnCamera);
            Console.WriteLine("List of pictures to be moved");
            foreach (string picture in pictures)
            {
                Console.WriteLine(picture.ToString());
            }

            string ORDERTYPE = getOrderType(); //-- > get from user input

            string orderNumber = getOrderNumber(ORDERTYPE);
            Directory.CreateDirectory(imageDestination + "\\" + ORDERTYPE + orderNumber);

            string FAXDATA = "";
            string BIN = "";
            string MLTSTORE = "";
            string DATA = "";
            string NOTE2 = "";
            string TAG = "";
            string REFERENCE2 = "";


            DateTime DATE = DateTime.Today; // --> M / D / YYYY

            DateTime dt = DateTime.Now;
            string TIME = String.Format("{0:HH:mm}", dt); //-- > 24hr

            int previousPicCount = getPhotoNo(orderNumber, ORDERTYPE);
            int numberOfPicsAdded = pictures.Count();
            int currentNumberOfPictures = previousPicCount + numberOfPicsAdded;
            decimal currentNumberOfPicturesAsDecimal = Convert.ToDecimal(currentNumberOfPictures);

            string NAME = getName(orderNumber, ORDERTYPE); //-- > get from work order

            string ACCOUNT = getAccount(orderNumber, ORDERTYPE); //-- > get from work order

            int ID = getIdNumber() + 1;

            Random random = new Random();
            int iteratorForOrders = 1;
            
            foreach (string nameOfFile in pictures)
            {
                Console.WriteLine("Image " + iteratorForOrders);
                string FILENAME = Path.GetFileName(nameOfFile).ToString();
                string extension = Path.GetExtension(nameOfFile).ToString();

                string DOCTYPE = extension.TrimStart('.');
                Console.WriteLine("DOCTYPE = " + DOCTYPE.ToUpper());

                string REFERENCE = buildReference(previousPicCount, iteratorForOrders, orderNumber);
                string NOTE = REFERENCE;

                Console.WriteLine("Moving File:");
                Console.WriteLine("ORDERTYPE = " + ORDERTYPE);
                Console.WriteLine("orderNumber = " + orderNumber);
                Console.WriteLine("NOTE = " + NOTE);
                Console.WriteLine("DATA = " + DATA);
                Console.WriteLine("FAXDATA = " + FAXDATA);
                Console.WriteLine("BIN = " + BIN);
                Console.WriteLine("MLTSTORE = " + MLTSTORE);
                Console.WriteLine("NOTE2 = " + NOTE2);
                Console.WriteLine("TAG = " + TAG);
                Console.WriteLine("REFERENCE = " + REFERENCE);
                Console.WriteLine("REFERENCE2 = " + REFERENCE2);
                Console.WriteLine("DATE = " + DATE);
                Console.WriteLine("TIME = " + TIME);
                Console.WriteLine("NAME = " + NAME);
                Console.WriteLine("ACCOUNT = " + ACCOUNT);
                Console.WriteLine("ID = " + ID);

                System.IO.File.Move(nameOfFile, String.Format("{0}\\{1}{2}\\{1}{2}-{3}-{4}{5}", imageDestination, ORDERTYPE, orderNumber, iteratorForOrders, getRandom(random), extension));

                Thread.Sleep(1000);

                Console.WriteLine(String.Format("{0}\\{1}{2}\\{1}{2}-{3}-{4}{5}", imageDestination, ORDERTYPE, orderNumber, iteratorForOrders, getRandom(random), extension));

                appendDb(NAME, ACCOUNT, ID.ToString(), DATE, TIME, DOCTYPE, REFERENCE, NOTE, DATA, FAXDATA, BIN, MLTSTORE, NOTE2, ORDERTYPE, FILENAME, TAG, REFERENCE2);

                ID++;
                iteratorForOrders++;
            }

            updatePhotoNo(orderNumber, ORDERTYPE, currentNumberOfPicturesAsDecimal);
            
            //NEED TO DO MORE RESEARCH BEFORE USING THIS, ERRORS OUT NOW
            //indexDb();

            Console.WriteLine("Process Complete");
            Console.WriteLine("Press Any Key To Continue");
            Console.ReadLine();
        }
    }
}
