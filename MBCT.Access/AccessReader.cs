using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Security.Cryptography;
using ADOX; //NOTE: Add Referance: Requires Microsoft ADO Ext. 2.8 for DDL and Security //https://www.microsoft.com/en-us/download/details.aspx?id=21995 //C:\Program Files\Common Files\System\ado\msadox.dll
using ADODB;
//NOTE: Add Referance: Extensions > ADODB


namespace MBCT.Access
{
    /// <summary>
    /// Hold onto lots of studio information to pass around while working with AccessTools.
    /// </summary>
    public class AccessReader : IDisposable
    {
        /// <summary>
        /// If database file is an older legacy file, use Jet, otherwise use Ace connection string.
        /// </summary>
        public bool Legacy;
        
        /// <summary>
        /// Only used if password protected.
        /// </summary>
        public string UserId;
        
        /// <summary>
        /// Only used if password protected.
        /// </summary>
        public string Password;
        
        /// <summary>
        /// Path to the Access database file.
        /// </summary>
        public string Path;
        
        /// <summary>
        /// Default table name if any.
        /// </summary>
        public string TableName;

        /// <summary>
        /// Initialize AccessReader with default values.
        /// </summary>
        public AccessReader()
        {
            ClearAttributes();
        }

        /// <summary>
        /// Initialize AccessReader with initial values.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="tablename"></param>
        /// <param name="password"></param>
        /// <param name="legacy"></param>
        /// <param name="userid"></param>
        public AccessReader(string path, string tablename = null, string password = null, bool legacy = false, string userid = null)
        {
            SetAttributes(path, tablename, password, legacy, userid);
        }

        //private calls
        private void SetAttributes(string path, string tablename, string password, bool legacy, string userid)
        {
            Legacy = legacy;
            UserId = userid;
            Password = password;
            Path = path;
            TableName = tablename;
        }
        private void ClearAttributes()
        {
            Legacy = false;
            UserId = null;
            Password = null;
            Path = null;
            TableName = null;
        }

        //public calls
        /// <summary>
        /// Returns a DataTable with results from provided query.
        /// </summary>
        /// <param name="strQuery"></param>
        /// <returns></returns>
        public DataTable Read(string strQuery = null)
        {
            if (!String.IsNullOrWhiteSpace(strQuery))
            {
                return AccessTools.CommandReader(Path, strQuery, Password, Legacy, UserId);
            }
            else
            {
                return AccessTools.Reader(Path, TableName, Password, Legacy, UserId);
            }
        }
        
        /// <summary>
        /// Update database file based on provided query.
        /// </summary>
        /// <param name="strQuery"></param>
        /// <returns></returns>
        public int Write(string strQuery = null)
        {
            return AccessTools.CommandWriter(Path, strQuery, Password, Legacy, UserId);
        }

        /// <summary>
        /// Adds provided table to the database file.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="pb1"></param>
        /// <returns></returns>
        public int Write(DataTable dt, ProgressBar pb1 = null)
        {
            AccessTools.Writer(Path, dt, Password, pb1, Legacy, UserId);
            return 0;
        }
        
        /// <summary>
        /// Adds provided dataset tables to the database file.
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="pb1"></param>
        /// <returns></returns>
        public int Write(DataSet ds, ProgressBar pb1 = null)
        {
            AccessTools.Writer(Path, ds, Password, pb1, Legacy, UserId);
            return 0;
        }
        
        /// <summary>
        /// Returns true if file extension matchs common file formats.
        /// </summary>
        /// <returns></returns>
        public bool IsAccess()
        {
            return AccessTools.IsAccess(Path);
        }
        
        /// <summary>
        /// Returns true if table exists. 
        /// </summary>
        /// <param name="tablename"></param>
        /// <returns></returns>
        public bool TableExists(string tablename = null)
        {
            tablename = tablename == null ? TableName : tablename;
            return AccessTools.TableExists(Path, tablename);
        }
        
        /// <summary>
        /// Returns a list of tables.
        /// </summary>
        /// <returns></returns>
        public DataTable GetTableList()
        {
            return AccessTools.GetTableList(Path, Password);
        }
        
        /// <summary>
        /// Returns a list of the columns names in provided table name.
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public DataTable GetTableColumnList(string tableName = null)
        {
            if (String.IsNullOrWhiteSpace(tableName))
            {
                tableName = TableName;
            }

            return AccessTools.GetTableColumnList(Path, tableName, Password);
        }
        
        /// <summary>
        /// Update given column name with new name.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <param name="newColumnName"></param>
        /// <returns></returns>
        public bool RenameColumn(string filePath, string tableName, string columnName, string newColumnName)
        {
            return AccessTools.RenameColumn(Path, TableName, columnName, newColumnName);
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        /// <summary>
        /// Dispose
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                ClearAttributes();
            }
            // free native resources if there are any.
        }
    }

}
