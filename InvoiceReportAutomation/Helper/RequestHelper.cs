using Dapper;
using InvoiceReportAutomation.Model;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using Helper;
using CrystalDecisions.Shared;

namespace InvoiceReportAutomation.Helper
{
    public class RequestHelper
    {
        public static string LasteErrorMessage { get; set; } = string.Empty;
        public static string DbConnectString_Midware { get; set; }  = ConfigurationManager.ConnectionStrings["MWconnectionString"].ConnectionString;
        public static bool IsBusy { get; set; } = false;

        static readonly string _fileId = "Invoice";

        public static void ExecuteRequest()
        {
            try
            {
                if (IsBusy) return;
                IsBusy = true;

                var sql = @"SELECT * FROM SAPInvoice 
                            WHERE IsCreated = 0 and isTry < 3";

                var conn = new SqlConnection(DbConnectString_Midware);
                var requests = conn.Query<RequestModel>(sql).ToList();

                if (requests == null)
                {
                    IsBusy = false;
                    return;
                }
                if (requests.Count == 0)
                {
                    IsBusy = false;
                    return;
                }

                HandleRequest(requests);
                IsBusy = false;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
                IsBusy = false;
            }
        }

        static void HandleRequest(List<RequestModel> requests)
        {
            try
            {
                if (requests == null) return;
                if (requests.Count == 0) return;
                requests.ForEach(x =>
                {
                    StartGenerateReport(x);
                });

            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }

        static void GenerateReport(RequestModel request)
        {
            try
            {
                var filepath = GetReportFilePath();
                var crConf = GetConnection(request);
                string pathcombined = Path.Combine(filepath.LayoutPath, filepath.LayoutName);

                string reportName = $"Invoice_{request.CompanyId}_{request.InvoiceDocEntry}.pdf";
                string destinationpath = Path.Combine(filepath.DestinationPath, reportName);

                var rpt = new CrystalRpt(pathcombined, crConf);
                //rpt.ServerName = crConf.Server;
                //rpt.DatabaseName = crConf.CompanyDB;
                //rpt.UserID = crConf.DbUserName;
                //rpt.Password = crConf.DbPassword;
                rpt.SetParamValue("DocKey@", request.InvoiceDocEntry);

                rpt.Connect();
                rpt.SaveAs(destinationpath, ExportFormatType.PortableDocFormat);

                UpdateSuccess(request.Id, reportName);
                Log($"{request.InvoiceDocEntry} success to generate.\n");
            }
            catch (Exception e)
            {
                Log($"{request.InvoiceDocEntry} fail.\n{e.Message}\n{ e.StackTrace}");
                UpdateFailed(request.Id);
            }
        }

        static void StartGenerateReport(RequestModel request)
        {
            try
            {
                var filepath = GetReportFilePath();
                var crConf = GetConnection(request);

                string pathcombined = Path.Combine(filepath.LayoutPath, filepath.LayoutName);

                string reportName = $"Invoice_{request.CompanyId}_{request.InvoiceDocEntry}.pdf";
                string destinationpath = Path.Combine(filepath.DestinationPath,reportName);

                var report = new CrystalReport(pathcombined, crConf);

                report.SetParameterValue("DocKey@", request.InvoiceDocEntry);

                report.SaveAsPdf(destinationpath);
                report.Close();

                UpdateSuccess(request.Id, reportName);
                Log($"{request.InvoiceDocEntry} success to generate.\n");
            }
            catch (Exception e)
            {
                Log($"{request.InvoiceDocEntry} fail.\n{e.Message}\n{ e.StackTrace}");
                UpdateFailed(request.Id);
            }
        }

        static RptLayoutPath GetReportFilePath()
        {
            try
            {
                var conn = new SqlConnection(DbConnectString_Midware);

                var query = "SELECT * FROM RptLayoutPath WHERE FileId = @FileId";

                var crCon = conn.Query<RptLayoutPath>(query, new { FileId = _fileId }).FirstOrDefault();

                if (crCon == null) throw new Exception($"Failed to get file path.");

                return crCon;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
                return null;
            }
        }

        static CrConfiguration GetConnection(RequestModel request)
        {
            try
            {
                if (string.IsNullOrEmpty(request.CompanyId)) throw new Exception($"CompanyId is empty. [ConnetionString-[{request.Id}]");
                var conn = new SqlConnection(DbConnectString_Midware);

                var query = "SELECT CompanyID, Server, DbUserName, DbPassword, CompanyDB FROM DBCommon WHERE CompanyID = @CompanyId";

                var crCon = conn.Query<CrConfiguration>(query, new { CompanyId = request.CompanyId }).FirstOrDefault();

                if (crCon == null) throw new Exception($"Failed to get connection string. [{request.Id}]");

                return crCon;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
                return null;
            }
        }

        static void UpdateSuccess(int DocId, string destination)
        {
            try
            {
                var conn = new SqlConnection(DbConnectString_Midware);

                var query = "UPDATE SAPInvoice set IsCreated = 1, FilePath = @DestinationPath, CreatedDate = GETDATE() WHERE Id = @DocId";

                conn.Execute(query, new { DocId = DocId, DestinationPath = destination });

            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }

        static void UpdateFailed(int DocId)
        {
            try
            {
                var conn = new SqlConnection(DbConnectString_Midware);

                var query = "UPDATE SAPInvoice set IsTry = IsTry + 1 WHERE Id = @DocId";

                conn.Execute(query, new { DocId = DocId });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }

        static void Log(string message)
        {
            Program.filelogger.Log(message.ToString());
        }
    }
}
