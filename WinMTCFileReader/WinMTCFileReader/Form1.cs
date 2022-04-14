using Npgsql;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WinMTCFileReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExecuteService();
        }


        #region Variables

        Uri _URLFTP = new Uri(@"ftp://" + ConfigurationManager.AppSettings["FTPHost"] + ":" + ConfigurationManager.AppSettings["FTPPort"] + ConfigurationManager.AppSettings["FTPFolder"]);
        string _FTPUSER = ConfigurationManager.AppSettings["FTPUser"];
        string _FTPPASS = ConfigurationManager.AppSettings["FTPPass"];
        string _TEMPFOLDER = AppDomain.CurrentDomain.BaseDirectory + "/TEMP/";

        #endregion

        #region Log

        /// <summary>
        /// Log de mensajes.
        /// </summary>
        /// <param name="message">Mensaje a logguear.</param>
        public void LogMessage(string message)
        {
            message = string.Concat(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "\t", message);

            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\MTCREADER_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(message);
                }
            }
        }

        #endregion

        #region Métodos

        #region Método Inicial

        public void ExecuteService()
        {
            var fileList = new List<string>();
            try
            {
                fileList = ObtenerListaReportesFTP();
            }
            catch (Exception e)
            {
                LogMessage("[ObtenerListaReportesFTP] : " + e.Message);
            }

            foreach (var archivo in fileList)
            {
                var fileName = archivo.Substring(49);
                List<string> query;

                if (fileName != string.Empty && fileName.EndsWith(".xlsx"))
                {

                    LogMessage("Archivo " + fileName + " procesando...");

                    try
                    {
                        query = ObtenerQuery(fileName);

                        EjecutarQuery(query);

                        if (!CheckFtpIfFileExists(fileName))
                            MoverExcelProcesado(fileName);
                        else
                            MoverExcelProcesado(fileName, true);

                        EliminarArchivoTemporal(fileName);

                    }
                    catch (Exception ex)
                    {
                        LogMessage("ERROR: [" + ex.Message + "]");
                        LogMessage("ARCHIVO: [" + fileName + "]");
                    }
                }
            }


        }

        #endregion

        private List<string> ObtenerListaReportesFTP()
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(_URLFTP.AbsoluteUri);
            request.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            request.KeepAlive = false;
            request.UsePassive = true;
            request.UseBinary = true;

            request.Credentials = new NetworkCredential(_FTPUSER, _FTPPASS);
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string names = reader.ReadToEnd();

            reader.Close();
            response.Close();

            return names.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        private List<string> ObtenerQuery(string fileName)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(_URLFTP + fileName);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.KeepAlive = false;
            request.UsePassive = true;
            request.UseBinary = true;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(_FTPUSER, _FTPPASS);

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();


            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);

            var url3 = new Uri(_TEMPFOLDER + fileName);

            using (FileStream writer = new FileStream(url3.LocalPath, FileMode.Create))
            {

                long length = response.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[2048];

                readCount = responseStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    writer.Write(buffer, 0, readCount);
                    readCount = responseStream.Read(buffer, 0, bufferSize);
                }
            }

            reader.Close();
            response.Close();

            var excel = new FileInfo(_TEMPFOLDER + fileName);
            List<RawData> list = new List<RawData>();

            using (var package = new ExcelPackage(excel))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.First(x => x.Hidden == eWorkSheetHidden.Visible && x.Name.Trim().ToUpper() == "BCONTACT");
                var rawData = new RawData();

                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (worksheet.Cells[i, col].Value != null)
                        {
                            switch (col)
                            {
                                case 1:
                                    rawData.store_code = worksheet.Cells[i, col].Value.ToString();
                                    break;
                                case 2:
                                    rawData.assing_code = worksheet.Cells[i, col].Value.ToString();
                                    break;
                                case 3:
                                    if (Regex.IsMatch(worksheet.Cells[i, col].Value.ToString(), @"^\d+$"))
                                    {
                                        rawData.process_date = DateTime.FromOADate((double)worksheet.Cells[i, col].Value);
                                    }
                                    else
                                    {
                                        rawData.process_date = DateTime.Parse(worksheet.Cells[i, col].Value.ToString());
                                    }
                                    break;
                                case 4:
                                    if (Regex.IsMatch(worksheet.Cells[i, col].Value.ToString(), @"^\d+$"))
                                    {
                                        rawData.register_date = DateTime.FromOADate((double)worksheet.Cells[i, col].Value);
                                    }
                                    else
                                    {
                                        rawData.register_date = DateTime.Parse(worksheet.Cells[i, col].Value.ToString());
                                    }
                                    break;
                                case 5:
                                    if (Regex.IsMatch(worksheet.Cells[i, col].Value.ToString(), @"^\d+$"))
                                    {
                                        rawData.paid_date = DateTime.FromOADate((double)worksheet.Cells[i, col].Value);
                                    }
                                    else
                                    {
                                        rawData.paid_date = DateTime.Parse(worksheet.Cells[i, col].Value.ToString());
                                    }
                                    break;
                                case 6:
                                    rawData.concept = worksheet.Cells[i, col].Value.ToString();
                                    break;
                                case 7:
                                    rawData.total = decimal.Parse(worksheet.Cells[i, col].Value.ToString().Replace("$", ""));
                                    list.Add(rawData);
                                    rawData = new RawData();
                                    break;
                            }
                        }
                    }
                }
            }



            string withString = @"WITH DATA AS (
	                                        SELECT * FROM (
		                                        VALUES 
                                                    {0}
                                    ) AS ok (store_code,assing_code,process_date,register_date,	paid_date,concept,total)

                                 )
                                 
                                 INSERT INTO bnext_bi.bcontact_payment_details (store_code,assing_code,process_date,register_date,paid_date,concept,total)
                                 SELECT * FROM DATA as d WHERE NOT EXISTS(
	                                SELECT 1 FROM bnext_bi.bcontact_payment_details WHERE assing_code = d.assing_code AND concept = d.concept AND d.total = total AND store_code = d.store_code and d.paid_date = paid_date
                                 ) AND COALESCE(concept, '') <> '' AND COALESCE(store_code, '') <> '';

                                 WITH DATA AS (
	                                        SELECT * FROM (
		                                        VALUES 
                                                    {0}
                                    ) AS ok (store_code,assing_code,process_date,register_date,	paid_date,concept,total)

                                 )
                                 INSERT INTO bnext_bi.bcontact_payment_details_historic (store_code,assing_code,process_date,register_date,paid_date,concept,total)
                                 SELECT * FROM DATA;";

            StringBuilder valuesString = new StringBuilder();
            int total = 0;
            int totalMal = 0;
            List<string> lstResult = new List<string>();

            foreach (var row in list)
            {
                if (!string.IsNullOrWhiteSpace(row.store_code) && !string.IsNullOrWhiteSpace(row.concept))
                {
                    if (row.store_code.Trim().Length > 0 && row.concept.Trim().Length > 0)
                    {
                        string select = @" ('{0}','{1}','{2}'::DATE,'{3}'::DATE,'{4}'::DATE,'{5}',{6}) ";

                        select = string.Format(select, row.store_code, row.assing_code, row.process_date.ToString("yyyy-MM-dd"), row.register_date.ToString("yyyy-MM-dd"), row.paid_date.ToString("yyyy-MM-dd"), row.concept, row.total);

                        total++;

                        if (total < list.Count)
                            select += ",";

                        valuesString.AppendLine(select);

                        if (total % 10000 == 0)
                        {
                            int place = valuesString.ToString().LastIndexOf(',');
                            string query = valuesString.ToString().Remove(place, 1).Insert(place, "");
                            withString = string.Format(withString, query);

                            lstResult.Add(withString);
                            valuesString = new StringBuilder();
                            withString = @"WITH DATA AS (
	                                        SELECT * FROM (
		                                        VALUES 
                                                    {0}
                                    ) AS ok (store_code,assing_code,process_date,register_date,	paid_date,concept,total)

                                 )
                                 
                                 INSERT INTO bnext_bi.bcontact_payment_details (store_code,assing_code,process_date,register_date,paid_date,concept,total)
                                 SELECT * FROM DATA as d WHERE NOT EXISTS(
	                                SELECT 1 FROM bnext_bi.bcontact_payment_details WHERE assing_code = d.assing_code AND concept = d.concept AND d.total = total AND store_code = d.store_code and d.paid_date = paid_date
                                 ) AND COALESCE(concept, '') <> '' AND COALESCE(store_code, '') <> '';

                                 WITH DATA AS (
	                                        SELECT * FROM (
		                                        VALUES 
                                                    {0}
                                    ) AS ok (store_code,assing_code,process_date,register_date,	paid_date,concept,total)

                                 )
                                 INSERT INTO bnext_bi.bcontact_payment_details_historic (store_code,assing_code,process_date,register_date,paid_date,concept,total)
                                 SELECT * FROM DATA;";
                        }
                    }
                    else totalMal++;
                }
                else totalMal++;
            }

            if (!string.IsNullOrEmpty(valuesString.ToString()))
            {
                if (totalMal > 0)
                {
                    int place = valuesString.ToString().LastIndexOf(',');
                    string query = valuesString.ToString().Remove(place, 1).Insert(place, "");
                    withString = string.Format(withString, query);
                    lstResult.Add(withString);
                }
                else
                {
                    withString = string.Format(withString, valuesString);
                    lstResult.Add(withString);
                }
            }


            reader.Close();
            reader.Dispose();
            response.Close();

            return lstResult;
        }

        private void MoverExcelProcesado(string fileName, bool prefix = false)
        {
            var requestMove = (FtpWebRequest)WebRequest.Create(_URLFTP + fileName);
            requestMove.Method = WebRequestMethods.Ftp.Rename;
            requestMove.Credentials = new NetworkCredential("bnextmtc", "MtCc0nc1l14c10n");
            requestMove.RenameTo = "/Reporte Pago de Pólizas/Procesados/" + (prefix ? DateTime.Now.Ticks.ToString() + " - " : "") + fileName;
            requestMove.KeepAlive = false;
            requestMove.UsePassive = true;
            requestMove.UseBinary = true;
            requestMove.GetResponse().Close();

        }

        private bool CheckFtpIfFileExists(string fileName)
        {
            Uri serverFile = new Uri(_URLFTP + "Procesados/" + fileName);
            FtpWebRequest reqFtp = (FtpWebRequest)FtpWebRequest.Create(serverFile);
            reqFtp.Method = WebRequestMethods.Ftp.GetFileSize;
            reqFtp.Credentials = new NetworkCredential(_FTPUSER, _FTPPASS);
            try
            {
                FtpWebResponse response = (FtpWebResponse)reqFtp.GetResponse();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void EliminarArchivoTemporal(string fileName)
        {
            if (File.Exists(Path.Combine(_TEMPFOLDER, fileName)))
            {
                File.Delete(Path.Combine(_TEMPFOLDER, fileName));
            }
        }

        private void EjecutarQuery(List<string> queryList)
        {
            var totalRegistros = 0;
            foreach (string query in queryList)
            {
                using (var cn = new NpgsqlConnection(ConfigurationManager.AppSettings["DBLink"]))
                {
                    try
                    {
                        cn.Open();

                        using (var cmd = new NpgsqlCommand(query, cn))
                        {
                            cmd.CommandTimeout = 2400;
                            totalRegistros += cmd.ExecuteNonQuery();
                        }

                        cn.Close();
                    }
                    catch (NpgsqlException e)
                    {
                        cn.Close();
                        LogMessage("[EjecutarQuery] : " + e.Message + " - [Query] : " + query);
                    }
                }
            }
            LogMessage("Total de registros afectados : " + totalRegistros);
        }

        #endregion

    }
}
