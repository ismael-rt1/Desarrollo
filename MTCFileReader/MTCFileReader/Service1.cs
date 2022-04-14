using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MTCFileReader
{
    public partial class Service1 : ServiceBase
    {
        private bool _processRunning;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            LogMessage("Servicio iniciado.");

            Task.Factory.StartNew(() => ProcessFtpFiles());

            try
            {

                Timer _timer = new Timer(3600000);
                _timer.Elapsed += timer_Elapsed;
                _timer.Enabled = true;
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("Fallo en timer. Message: {0}. StackTrace: {1}", ex.Message, ex.StackTrace));
            }
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_processRunning)
            {
                return;
            }

            ProcessFtpFiles();
        }

        protected override void OnStop()
        {
            LogMessage("Servicio detenido.");
        }


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


        private void ProcessFtpFiles()
        {
            _processRunning = true;

            LogMessage("Iniciando proceso");

            var handler = new MTCReader();

            try
            {
                if (DateTime.Now.Hour >= 8 || DateTime.Now.Hour <= 20)
                {
                    handler.ExecuteService();
                }
            }
            catch (Exception e)
            {
                LogMessage(e.Message);
            }

            LogMessage("Terminando proceso");
            _processRunning = false;
        }

    }
}
