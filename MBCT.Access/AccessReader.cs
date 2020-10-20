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
    /// Hold onto lots of studio information to pass around
    /// </summary>
    public class AccessReader : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        public bool Legacy;
        /// <summary>
        /// 
        /// </summary>
        public string UserId;
        /// <summary>
        /// 
        /// </summary>
        public string Password;
        /// <summary>
        /// 
        /// </summary>
        public string Path;
        /// <summary>
        /// 
        /// </summary>
        public string TableName;

        /// <summary>
        /// 
        /// </summary>
        public AccessReader()
        {
            ClearAttributes();
        }

        /// <summary>
        /// 
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
        /// 
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
        /// 
        /// </summary>
        /// <param name="strQuery"></param>
        /// <returns></returns>
        public int Write(string strQuery = null)
        {
            return AccessTools.CommandWriter(Path, strQuery, Password, Legacy, UserId);
        }

        /// <summary>
        /// 
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
        /// 
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
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsAccess()
        {
            return AccessTools.IsAccess(Path);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tablename"></param>
        /// <returns></returns>
        public bool TableExists(string tablename = null)
        {
            tablename = tablename == null ? TableName : tablename;
            return AccessTools.TableExists(Path, tablename);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetTableList()
        {
            return AccessTools.GetTableList(Path, Password);
        }
        /// <summary>
        /// 
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
        /// 
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

        //Dispose
        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// 
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
