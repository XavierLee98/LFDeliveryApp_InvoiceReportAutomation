using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using InvoiceReportAutomation.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    class CrystalRpt : IDisposable
    {
        /// <summary>
        /// Default value: Current Server SAP is connected to
        /// </summary>
        public string ServerName
        {
            get => conn.ServerName;
            set => conn.ServerName = value;
        }

        /// <summary>
        /// Default value: Current database SAP is connected to
        /// </summary>
        public string DatabaseName
        {
            get => conn.DatabaseName;
            set => conn.DatabaseName = value;
        }

        public string UserID
        {
            get => conn.UserID;
            set => conn.UserID = value;
        }

        public string Password
        {
            get => conn.Password;
            set => conn.Password = value;
        }

        ReportDocument cr = new ReportDocument();
        ConnectionInfo conn = new ConnectionInfo();
        bool disposed = false;

        public CrystalRpt(string filepath, CrConfiguration CrConn)
        {
            cr.Load(filepath);
            ServerName = CrConn.Server;
            DatabaseName = CrConn.CompanyDB;
            UserID = CrConn.DbUserName;
            Password = CrConn.DbPassword;
        }

        ~CrystalRpt()
        {
            Dispose();
        }

        public void SaveAs(string filepath, ExportFormatType format) => cr.ExportToDisk(format, filepath);

        public void Connect() => cr.Database.Tables.Cast<Table>().ToList().ForEach(table => table.LogOnInfo.ConnectionInfo = conn);

        public void SetParamValue(string name, object value) => cr.SetParameterValue(name, value);

        public void Dispose()
        {
            lock (this)
            {
                if (disposed) return;

                disposed = true;
            }

            cr.Close();
            cr.Dispose();
        }
    }
}