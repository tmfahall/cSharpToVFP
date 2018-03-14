using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vfpTypeFinder
{
    class Program
    {
        static void Main(string[] args)
        {
            OleDbConnection connection = new OleDbConnection(
            "Provider=VFPOLEDB.1;Data Source=C:\\Users\\Administrator\\Desktop\\vision\\WO.DBF");

            connection.Open();
            DataTable tables = connection.GetSchema(OleDbMetaDataCollectionNames.Tables);

            foreach (DataRow rowTables in tables.Rows)
            {
                if (rowTables["table_name"].ToString().ToLower() == "wo")
                {
                    Console.WriteLine(rowTables["table_name"].ToString());
                }

                DataTable columns = connection.GetSchema(OleDbMetaDataCollectionNames.Columns, new String[] { null, null, rowTables["table_name"].ToString(), null });

                foreach (DataRow rowColumns in columns.Rows)
                {
                    if (rowColumns["table_name"].ToString().ToLower() == "wo")
                    {
                        Console.WriteLine(
                            rowTables["table_name"].ToString() + "." +
                            rowColumns["column_name"].ToString() + " = " +
                            rowColumns["data_type"].ToString()
                            );
                    }

                }
            }

            Console.ReadLine();
        }
    }
}

//adEmpty	    0	This data type indicates that no value was specified(DBTYPE_EMPTY).
//adSmallInt	2	This data type indicates a 2-byte (16-bit) signed integer(DBTYPE_I2).
//adInteger	    3	This data type indicates a 4-byte (32bit) signed integer(DBTYPE_I4).
//adSingle	    4	This data type indicates a 4-byte (32-bit) single-precision IEEE floating-point number                       (DBTYPE_R4).
//adDouble	    5	This data type indicates an 8-byte (64-bit) double-precision IEEE floating-point number                      (DBTYPE_R8).
//adCurrency	6	A data type indicates a currency value(DBTYPE_CY). Currency is a fixed-point number with                     four digits to the right of the decimal point.It is stored in an 8 - byte signed integer                     scaled by 10, 000.This data type is not supported by the Microsoft® OLE DB Provider for                      AS / 400 and VSAM or the Microsoft OLE DB Provider for DB2.
//adDate        7   This data type indicates a date value stored as a Double, the whole part of which is the                     number of days since December 30, 1899, and the fractional part of which is the fraction of                  a day.This data type is not supported by the OLE DB Provider for AS / 400 and VSAM or the                    OLE DB Provider for DB2.
//adBSTR        8   This data type indicates a null - terminated Unicode character string(DBTYPE_BSTR).This                      data type is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB                        Provider for DB2.
//adIDispatch   9   This data type indicates a pointer to an IDispatch interface on an OLE object                                (DBTYPE_IDISPATCH). This data type is not supported by the OLE DB Provider for AS/400 and                    VSAM or the OLE DB Provider for DB2.
//adError	    10	This data type indicates a 32-bit error code(DBTYPE_ERROR). This data type is not supported                  by the OLE DB Provider for AS/400 and VSAM or the OLE DB Provider for DB2.
//adBoolean	    11	This data type indicates a Boolean value(DBTYPE_BOOL). This data type is not supported by                    the OLE DB Provider for AS/400 and VSAM.
//adVariant     12	This data type indicates an Automation variant (DBTYPE_VARIANT). This data type is not                       supported by the OLE DB Provider for AS/400 and VSAM or the OLE DB Provider for DB2.
//adIUnknown    13	This data type indicates a pointer to an IUnknown interface on an OLE object                                 (DBTYPE_IUNKNOWN). This data type is not supported by the OLE DB Provider for AS/400 and                     VSAM or the OLE DB Provider for DB2.
//adDecimal	    14	This data type indicates numeric data with a fixed precision and scale(DBTYPE_DECIMAL).
//adTinyInt     16  This data type indicates a single - byte(8 - bit) signed integer(DBTYPE_I1).This data type                   is not supported by the OLE DB Provider.
//adUnsignedTinyInt   17  This data type indicates a single - byte(8 - bit) unsigned integer(DBTYPE_UI1).This                          data type is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB                        Provider for DB2.
//adUnsignedSmallInt  18  This data type indicates a 2 - byte(16 - bit) unsigned integer(DBTYPE_UI2).This data                         type is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB                             Provider for DB2.
//adUnsignedInt 19  This data type indicates a 4 - byte(32 - bit) unsigned integer(DBTYPE_UI4).This data type                    is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB Provider for                     DB2.
//adBigInt      20  This data type indicates an 8 - byte(64 - bit) signed integer(DBTYPE_I8).This data type is                   not supported by the OLE DB Provider for AS / 400 and VSAM.
//adUnsignedBigInt    21  This data type indicates an 8 - byte(64 - bit) unsigned integer(DBTYPE_UI8).This data                        type is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB                             Provider for DB2.
//adGUID        72  This data type indicates a globally unique identifier or GUID(DBTYPE_GUID).This data type                    is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB Provider for                     DB2.
//adBinary      128 This data type indicates fixed-length binary data(DBTYPE_BYTES).
//adChar        129 This data type indicates a character string value(DBTYPE_STR).
//adWChar       130 This data type indicates a null - terminated Unicode character string(DBTYPE_WSTR).This                      data type is not supported by the OLE DB Provider for AS / 400 and VSAM or the OLE DB                        Provider for DB2.
//adNumeric     131 This data type indicates numeric data where the precision and scale are exactly as                           specified (DBTYPE_NUMERIC).
//adUserDefined 132 This data type indicates user - defined data(DBTYPE_UDT).This data type is not supported by                  the OLE DB Provider for AS / 400 and VSAM or the OLE DB Provider for DB2.
//adDBDate      133 This data type indicates an OLE DB date structure(DBTYPE_DATE).
//adDBTime      134 This data type indicates an OLE DB time structure(DBTYPE_TIME).
//adDBTimeStamp 135 This data type indicates an OLE DB timestamp structure(DBTYPE_TIMESTAMP).
//adVarChar     200 This data type indicates variable - length character data(DBTYPE_STR).
//adLongVarChar 201 This data type indicates a long string value.
//adVarWChar    202 This data type indicates a Unicode string value.This data type is not supported by the OLE                   DB Provider for AS / 400 and VSAM or the OLE DB Provider for DB2.
//adLongVarWChar    203 This data type indicates a long Unicode string value.This data type is not supported by                      the OLE DB Provider for AS / 400 and VSAM or the OLE DB Provider for DB2.
//adVarBinary   204 This data type indicates variable - length binary data(DBTYPE_BYTES).
//adLongVarBinary   205 This data type indicates a long binary value.