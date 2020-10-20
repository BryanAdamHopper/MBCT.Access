using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//Added Referance: System.Data; System.Windows.Forms;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Security.Cryptography;
using ADOX; //Add Referance: Requires Microsoft ADO Ext. 2.8 for DDL and Security //https://www.microsoft.com/en-us/download/details.aspx?id=21995 //C:\Program Files\Common Files\System\ado\msadox.dll
using ADODB;

//Add Referance: Extensions > ADODB

//using System.Data.Odbc;

namespace MBCT.Access
{

    /// <summary>
    /// 
    /// </summary>
    public class AccessTools
    {
        /// <summary>
        /// 
        /// </summary>
        public static readonly string Type = "Access";
        /// <summary>
        /// 
        /// </summary>
        public static readonly List<string> Types = new List<string>
        {
            "accdb",
            "mdb",
            "adp",
            "mda",
            "accda",
            "mde",
            "accde",
            "ade"
        };
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<string> OpenFileDialogueFilter()
        {
            var types = Types.Aggregate("", (current, type) => current + $"*.{type};");
            types += "|All Files|*.*";
            return new List<string> { $"{Type}", $"{Type} Files|{types}" };
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static OpenFileDialog OpenFileDiag()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            List<string> types = OpenFileDialogueFilter();

            ofd.Title = $"Open {types[0]} File";
            ofd.Filter = $"{types[1]}";
            ofd.FileName = null;

            return ofd;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool CreateNewAccessDatabase(string fileName)
        {
            bool result = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            //ADOX.Table table = new ADOX.Table();

            //Create the table and it's fields. 
            //table.Name = "Table1";
            //table.Columns.Append("Field1");
            //table.Columns.Append("Field2");

            try
            {
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
                //cat.Tables.Append(table);

                //Now Close the database
                ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
                if (con != null)
                    con.Close();

                result = true;
            }
            catch (Exception)
            {
                result = false;
            }
            cat = null;
            return result;
        }

        /// <summary>
        /// Replaces common issue text characters with its unicode equivalent.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static string Sanatize(string text)
        {
            //text = text.Replace("" + (char)160, "").Replace("'", @"\u0027").Replace("" + (char)34, @"\u0022").Replace("" + (char)10, @"\n").Replace("" + (char)13, @"\r").Trim();
            text = text.Replace("" + (char)160, "").Replace("'", "\'\'").Replace("" + (char)34, "\"\"").Replace("" + (char)10, "\n").Replace("" + (char)13, "\r").Trim();
            return text;
        }

        /// <summary>
        /// Returns true if file extension matchs common file formats.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IsAccess(string path)
        {
            string ext = Path.GetExtension(path);
            string filesTypes = "*.accdb; *.mdb; *.adp; *.mda; *.accda; *.mde; *.accde; *.ade";

            if (filesTypes.IndexOf(ext) > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        /// <summary>
        /// Returns true if table exists. 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static bool TableExists(string path, string tableName)
        {
            using (OleDbConnection conn = new OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", path)))
            {
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                return conn.GetSchema("Tables", new string[4] { null, null, tableName, "TABLE" }).Rows.Count > 0;
            }
        }

        /// <summary>
        /// Returns true if column exists.
        /// 
        /// OleDbSchemaGuid.Columns, new[] {TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME}
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static bool ColumnExists(string path, string tableName, string columnName)
        {
            using (OleDbConnection conn = new OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", path)))
            {
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                using (DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new[] {null, null, tableName, columnName}))
                {
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static int GetRowCount(string path, string table)
        {
            int rows = 0;
            string strSQL = "";
            string strText = "";

            try
            {
                strSQL = String.Format("select count(*) from [{0}]", table);
                strText = CommandReader(path, strSQL).Rows[0][0].ToString();
                rows = int.Parse(strText);
            }
            catch
            {
                rows = -1;
            }

            return rows;
        }

        /// <summary>
        /// Return list of tables in file.
        /// The object here represents the 1st 4 columns, the 4th one being TABLE_TYPE.
        /// TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, TABLE_GUID, DESCRIPTION, TABLE_PROPID, DATE_CREATED, DATE_MODIFIED
        /// </summary>
        /// <param name="path"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DataTable GetTableList(string path, string password = "")
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path, password)))
            {
                conn.ConnectionString = AccessTools.GetConnectionString(path, password);
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                /*The object here represents the 1st 4 columns, the 4th one being TABLE_TYPE.*/
                /*TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, TABLE_GUID, DESCRIPTION, TABLE_PROPID, DATE_CREATED, DATE_MODIFIED*/
                return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            }
        }

        /// <summary>
        /// Return list of columns in table.
        /// The object here represents the 1st 4 columns, the 3rd one being TABLE_NAME.
        /// TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tableName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DataTable GetTableColumnList(string path, string tableName, string password = "")
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path, password)))
            {
                conn.ConnectionString = AccessTools.GetConnectionString(path, password);
                if (conn.State != ConnectionState.Open) { conn.Open(); }
                /*The object here represents the 1st 4 columns, the 3rd one being TABLE_NAME.*/
                /*TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME*/
                return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new[] { null, null, tableName});
            }
        }

        /// <summary>
        /// Creates connection string based on file type and parameters provided by user.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="password"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        public static string GetConnectionString(string path, string password = "", bool isLegacy = false, string userID = "admin")
        {
            //string ext = Path.GetExtension(path);
            string connectionString = "";

            if (isLegacy == true)
            {
                connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", path);
            }
            else
            {
                connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", path);
            }

            if (String.IsNullOrEmpty(password))
            {
                if (isLegacy == true)
                {
                    connectionString += String.Format("User Id={0};Password={1};", userID, password);
                }
                else
                {
                    connectionString += "Persist Security Info=False;";
                }
            }
            else
            {
                connectionString += String.Format("Jet OLEDB:Database Password={0};", password);
            }

            return connectionString;
        }

        /// <summary>
        /// Read from Access.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tableName"></param>
        /// <param name="password"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        public static DataTable Reader(string path, string tableName, string password = "", bool isLegacy = false, string userID = "admin")
        {
            //OleDbConnection conn = new OleDbConnection(); { conn.ConnectionString = getConnectionString(path, password, isLegacy, userID); }
            //OleDbCommand comm = new OleDbCommand(); { comm.Connection = conn; }
            ////string connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", path);
            ////connString = getConnectionString(path, password);
            ////conn.ConnectionString = getConnectionString(path, password);
            ////comm.Connection = conn;
            //comm.CommandText = String.Format("SELECT * FROM [{0}]", tableName);

            //DataTable dt = new DataTable();
            //if (conn.State != ConnectionState.Open) { conn.Open(); }
            //dt.Load(comm.ExecuteReader());
            //if (conn.State != ConnectionState.Closed) { conn.Close(); }

            //return dt;

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path, password, isLegacy, userID)))
            {
                using (OleDbCommand comm = new OleDbCommand(String.Format("SELECT * FROM [{0}]", tableName), conn))
                {
                    using (DataTable dt = new DataTable())
                    {
                        if (conn.State != ConnectionState.Open) { conn.Open(); }
                        dt.Load(comm.ExecuteReader());
                        return dt;
                    }
                }
            }


        }

        /// <summary>
        /// Access Writer. 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="strSQL"></param>
        /// <param name="password"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        public static int CommandWriter(string path, string strSQL, string password = "", bool isLegacy = false, string userID = "admin")
        {
            int results;

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path, password, isLegacy, userID)))
            {
                using (OleDbCommand comm = new OleDbCommand(strSQL, conn))
                {
                    if (conn.State != ConnectionState.Open) { conn.Open(); }
                    results = comm.ExecuteNonQuery();
                }
            }

            return results;
        }

        /// <summary>
        /// Access Reader.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="strSQL"></param>
        /// <param name="password"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        public static DataTable CommandReader(string path, string strSQL, string password = "", bool isLegacy = false, string userID = "admin")
        {
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(path, password, isLegacy, userID)))
            {
                using (OleDbCommand comm = new OleDbCommand(strSQL, conn))
                {
                    using (DataTable dt = new DataTable())
                    {
                        if (conn.State != ConnectionState.Open) { conn.Open(); }
                        dt.Load(comm.ExecuteReader());
                        return dt;
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <param name="newColumnName"></param>
        /// <returns></returns>
        public static bool RenameColumn(string filePath, string tableName, string columnName, string newColumnName)
        {
            bool pass = false;
            try
            {
                string q = String.Format(@"ALTER TABLE [{0}] RENAME COLUMN [{1}] TO [{2}]", tableName, columnName, newColumnName);
                AccessTools.CommandWriter(filePath, q);
                pass = true;
            }
            catch
            {
                pass = false;
            }
            return pass;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dt"></param>
        /// <param name="password"></param>
        /// <param name="pb1"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        public static void Writer(string path, DataTable dt, string password = null, ProgressBar pb1 = null, bool isLegacy = false, string userID = "admin")
        {
            _CreateTable(path, dt, password, isLegacy, userID); //Deletes table if exists, then creates new table.
            _PopulateTable(path, dt, password, pb1, isLegacy, userID);
            //_SanitizeCellData(path, dt, password, pb1, isLegacy, userID);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="ds"></param>
        /// <param name="password"></param>
        /// <param name="pb1"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        public static void Writer(string path, DataSet ds, string password = null, ProgressBar pb1 = null, bool isLegacy = false, string userID = "admin")
        {
            for (int x = 0; x < ds.Tables.Count; x++)
            {
                Writer(path, ds.Tables[0], password, pb1, isLegacy, userID);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="password"></param>
        /// <param name="isLegacy"></param>
        /// <param name="userID"></param>
        /// <returns></returns>
        private static OleDbCommand AccessDB(string path, string password = null, bool isLegacy = false, string userID = "admin")
        {
            OleDbConnection conn = new OleDbConnection();
            { conn.ConnectionString = GetConnectionString(path, password, isLegacy, userID); }
            OleDbCommand comm = new OleDbCommand();
            { comm.Connection = conn; }

            conn.Open();

            return comm;
        }
        private static string GetAccessType(Type type)
        {
            //if(!type.Contains("System."))
            //{
            //    type = String.Format("System.{0}", type);
            //}

            string strType = type.ToString().Replace("System.", "");

            return GetAccessType(strType);    //System.Type.GetType(type)
        }
        private static string GetAccessType(string type)
        {
            string returnType = "";
            string strType = type;  //type.ToString().Replace("System.","");


            #region DataTable Types
            /*
                Boolean
                Byte
                Byte[]
                Char
                DateTime
                Decimal
                Double
                Guid
                Int16
                Int32
                Int64
                SByte
                Single
                String
                TimeSpan
                UInt16
                UInt32
                UInt64
            */
            #endregion

            switch (strType)
            {
                case "Boolean":
                    {
                        returnType = "YesNo";
                        break;
                    }
                case "Byte":
                    {
                        returnType = "Byte";
                        break;
                    }
                case "Byte[]":
                    {
                        returnType = "OLEObject";
                        break;
                    }
                case "Char":
                    {
                        returnType = "Char";
                        break;
                    }
                case "DateTime":
                    {
                        returnType = "DateTime";
                        break;
                    }
                case "Decimal":
                    {
                        returnType = "Double";
                        break;
                    }
                case "Double":
                    {
                        returnType = "Double";
                        break;
                    }
                case "GUID":
                    {
                        returnType = "GUID";
                        break;
                    }
                case "Int16":
                    {
                        returnType = "Short";
                        break;
                    }
                case "Int32":
                    {
                        returnType = "Long";
                        break;
                    }
                case "Int64":
                    {
                        returnType = "Text(25)";
                        break;
                    }
                case "Sbyte":
                    {
                        returnType = "Short";
                        break;
                    }
                case "Single":
                    {
                        returnType = "Single";
                        break;
                    }
                case "String":
                    {
                        returnType = "Text(255)"; // or MEMO
                        break;
                    }
                case "StringShort":
                    {
                        returnType = "Text(255)"; // or MEMO
                        break;
                    }
                case "StringLong":
                    {
                        returnType = "Memo"; // or MEMO
                        break;
                    }
                case "TimeSpan":
                    {
                        returnType = "DateTime";
                        break;
                    }
                case "UInt16":
                    {
                        returnType = "Long";
                        break;
                    }
                case "UInt32":
                    {
                        returnType = "Double";
                        break;
                    }
                case "UInt64":
                    {
                        returnType = "Text(26)";
                        break;
                    }


            }

            return returnType;
        }
        private static bool _DeleteTable(string path, string tableName, string password = null, bool isLegacy = false, string userID = "admin")
        {
            if (TableExists(path, tableName) == true)
            {
                string strSQL = String.Format("\tDROP TABLE [{0}]", tableName);

                using (OleDbCommand comm = AccessDB(path, password, isLegacy, userID))
                {
                    comm.CommandText = strSQL;
                    comm.ExecuteNonQuery();
                }

                return true;
            }
            else
            {
                return false;
            }
        }
        private static void _CreateTable(string path, DataTable dt, string password = null, bool isLegacy = false, string userID = "admin")
        {
            string TABLE_NAME = dt.TableName;
            string strSQL = "";

            //drop table if exists
            _DeleteTable(path, TABLE_NAME, password, isLegacy, userID);

            //get column headers and datatypes (ID TEXT(255), ContactLog MEMO)
            string COLUMN_HEADERS = "";
            string COLUMN_HEADERS_TYPES = "";


            for (int col = 0; col < dt.Columns.Count; col++)
            {
                if (col != 0)
                {
                    COLUMN_HEADERS += ", ";
                    COLUMN_HEADERS_TYPES += ", ";
                }

                string wrappedHeader = $"[{dt.Columns[col].ColumnName.ToString()}]";
                COLUMN_HEADERS += wrappedHeader;
                COLUMN_HEADERS_TYPES += wrappedHeader;

                string colType = dt.Columns[col].DataType.ToString().Replace("System.", "");
                //MessageBox.Show(String.Format("colType: {0}", colType));

                //determine if each column is TEXT or MEMO based on MAX(Length) of column rowCells.
                if (colType == "String")
                {
                    colType = "StringShort";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        //look for a cell that contains more than 255 chars
                        //if found then StringLong, otherwise StringShort
                        //if (sanatize(dt.Rows[row][col].ToString()).Length > 255)
                        if (dt.Rows[row][col].ToString().Length > 255)
                        {
                            colType = "StringLong";
                            break;
                        }

                    }
                }

                //convert DataTable type to Access type
                colType = GetAccessType(colType);

                COLUMN_HEADERS_TYPES += String.Format(" {0}", colType);
            }


            strSQL = String.Format("CREATE TABLE [{0}] ({1})", TABLE_NAME, COLUMN_HEADERS_TYPES);
            using (OleDbCommand comm = AccessDB(path, password, isLegacy, userID))
            {
                comm.CommandText = strSQL;
                comm.ExecuteNonQuery();
            }

        }
        private static void _PopulateTable(string path, DataTable dt, string password = null, ProgressBar pb1 = null, bool isLegacy = false, string userID = "admin")
        {
            string TABLE_NAME = dt.TableName;

            string COLUMN_HEADERS = "";
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                if (col != 0) { COLUMN_HEADERS += ", "; }
                COLUMN_HEADERS += $"[{dt.Columns[col].ColumnName.ToString()}]";
            }

            // ##### POPULATE TABLE ###########################################################################
            using (OleDbCommand comm = AccessDB(path, password, isLegacy, userID))
            {
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    string strRowText = "";

                    for (int y = 0; y < dt.Columns.Count; y++)
                    {
                        if (y != 0) { strRowText += ", "; }
                        string cellData = Sanatize(dt.Rows[x][y].ToString());
                        if (dt.Columns[y].DataType.ToString().Replace("System.", "") == "Boolean")
                        {
                            cellData = cellData.ToUpper() == "True".ToUpper() ? "-1" : "0";
                        }

                        strRowText += string.Format("'{0}'", cellData);

                    }

                    string strSQL = string.Format("INSERT INTO [{0}] ({1}) VALUES ({2})", TABLE_NAME, COLUMN_HEADERS, strRowText);
                    try
                    {
                        if (String.IsNullOrWhiteSpace(strRowText.Replace(",", "").Replace("'", "")) == false)
                        {
                            comm.CommandText = strSQL;
                            comm.ExecuteNonQuery();
                        }

                    }
                    catch (Exception ex)
                    {
                        String.Format(ex.Message);
#if DEBUG
                        MessageBox.Show(String.Format("strSQL: \n{0}\n{1}", strSQL, ex.ToString()), String.Format("(DEBUG only message) Error: {0}", ex.Message), MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
                    }



                    if (pb1 != null) { pb1.Value = (x * 100) / dt.Rows.Count; Application.DoEvents(); }
                }

            }

            if (pb1 != null) { pb1.Value = 0; Application.DoEvents(); }

        }
        private static void _SanitizeCellData(string path, DataTable dt, string password = null, ProgressBar pb1 = null, bool isLegacy = false, string userID = "admin")
        {
            //string TABLE_NAME = dt.TableName;

            //using (OleDbCommand comm = AccessDB(path, password, isLegacy, userID))
            //{
            //    for (int c = 0; c < dt.Columns.Count; c++)
            //    {
            //        if (dt.Columns[c].DataType.ToString().Replace("System.", "") == "String")
            //        {
            //            string strSQL = String.Format("UPDATE [{0}] SET [{1}] = REPLACE([{1}], '{2}', '{3}') WHERE [{1}] IS NOT NULL", TABLE_NAME, dt.Columns[c].ColumnName.ToString(), @"\u0027", "''");
            //            comm.CommandText = strSQL;
            //            comm.ExecuteNonQuery();

            //            strSQL = String.Format("UPDATE [{0}] SET [{1}] = REPLACE([{1}], '{2}', '{3}') WHERE [{1}] IS NOT NULL", TABLE_NAME, dt.Columns[c].ColumnName.ToString(), @"\u0022", (char)34);
            //            comm.CommandText = strSQL;
            //            comm.ExecuteNonQuery();

            //            strSQL = String.Format("UPDATE [{0}] SET [{1}] = REPLACE([{1}], '{2}', '{3}') WHERE [{1}] IS NOT NULL", TABLE_NAME, dt.Columns[c].ColumnName.ToString(), @"\r", "" + (char)13);
            //            comm.CommandText = strSQL;
            //            comm.ExecuteNonQuery();

            //            strSQL = String.Format("UPDATE [{0}] SET [{1}] = REPLACE([{1}], '{2}', '{3}') WHERE [{1}] IS NOT NULL", TABLE_NAME, dt.Columns[c].ColumnName.ToString(), @"\n", "" + (char)10);
            //            comm.CommandText = strSQL;
            //            comm.ExecuteNonQuery();
            //        }
            //    }
            //}
        }
    }
}
