using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Net;

using System.Security.Policy;
using Renci.SshNet;
using Renci.SshNet.Sftp;


namespace Descomprimir
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            
        }

        private void button1_Click(object sender, EventArgs e)
        {


            //[*RECORRE TODOS LOS ARCHIVOS QUE SE ENCUENTREN EN LA RUTA *]
            DirectoryInfo dir = new DirectoryInfo(@"C:\Users\iruiz\Documents");

            foreach (var fi in dir.GetFiles())
            {
                Console.WriteLine(fi.Name);
                // MessageBox.Show(fi.Name);

            }


           //[*DESCOMPRIME ARCHIVO ESPECIFICADO QUE SE ENCUENTREN EN LA RUTA *]
            DirectoryInfo ZipdirInfo = new DirectoryInfo(@"C:\Users\iruiz\Documents");

            foreach (FileInfo zipFilesInfo in ZipdirInfo.GetFiles("BNEXT864200820.txt.gz"))
            {

                      unZip(zipFilesInfo);
                       MessageBox.Show("Se descomprime archivo :"+ zipFilesInfo);

            }
                 

            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\iruiz\Documents\");   
            //  string basePath = Environment.CurrentDirectory;

        }

       public void descomprimirArchivo()
        {
            string directorioComprimido = @"C:\Users\iruiz\Documents\BNEXT864200820.txt.gz";
            string directorioDestindo = @"C:\Users\iruiz\Documents";

            //Descomprimir
            System.IO.Compression.ZipFile.ExtractToDirectory(directorioComprimido, directorioDestindo);

        }



        public static void ZipFiles(FileInfo zipFilesInfo)

		{

			using (FileStream varFileStream = zipFilesInfo.OpenRead())

			{

				if ((File.GetAttributes(zipFilesInfo.FullName) & FileAttributes.Hidden) != FileAttributes.Hidden & zipFilesInfo.Extension != ".gz")

				{

					using (FileStream varOutFileStream =

						File.Create(zipFilesInfo.FullName + ".gz"))

					{

						using (GZipStream Zip = new GZipStream(varOutFileStream,

						CompressionMode.Compress))
						{

							varFileStream.CopyTo(Zip);

						}

					}

				}

			}

		}

		public static void unZip(FileInfo unzipFile)
		{
			using (FileStream zipFile = unzipFile.OpenRead())

			{

				string zipCurFile = unzipFile.FullName;
				string origZipFileName = zipCurFile.Remove(zipCurFile.Length - unzipFile.Extension.Length);

				using (FileStream outunzipFile = File.Create(origZipFileName))
				{
					using (GZipStream Decompress = new GZipStream(zipFile, CompressionMode.Decompress))

					{
						Decompress.CopyTo(outunzipFile);

						Console.WriteLine("Decompressed: {0}", unzipFile.Name);

					}

				}

			}

		}


        #region  Descomprime el archivo
        /// <summary>
        /// Descomprime el archivo en la carpeta especificada
        /// </summary>
        /// <param name = "srcZipFile"> archivo de origen comprimido </param>
        /// <param name = "destDir"> Carpeta de destino </param>
        /// <returns></returns>
        /// 

       

        public static void UnZipFile(string srcZipFile, string destDir)
        {
            // Lea el archivo comprimido (archivo zip) y prepárese para descomprimir
            ZipInputStream inputstream = new ZipInputStream(File.OpenRead(srcZipFile.Trim()));
            ZipEntry entry;
            string path = destDir;
            // Guardar la ruta del archivo descomprimido
            string rootDir = "";
            // El nombre de la primera subcarpeta del directorio raíz
            while ((entry = inputstream.GetNextEntry()) != null)
            {
                rootDir = Path.GetDirectoryName(entry.Name);
                // Obtenga el nombre de la subcarpeta de primer nivel en el directorio raíz
                if (rootDir.IndexOf("\\") >= 0)
                {
                    rootDir = rootDir.Substring(0, rootDir.IndexOf("\\") + 1);
                }
                string dir = Path.GetDirectoryName(entry.Name);
                // Obtenga el nombre de la subcarpeta en la subcarpeta de primer nivel en el directorio raíz
                string fileName = Path.GetFileName(entry.Name);
                // Nombre de archivo en el directorio raíz
                if (dir != "")
                {
                    // Cree una subcarpeta en el directorio raíz, sin limitar el nivel
                    if (!Directory.Exists(destDir + "\\" + dir))
                    {
                        path = destDir + "\\" + dir;
                        // Crea una carpeta en la ruta especificada
                        Directory.CreateDirectory(path);
                    }
                }
                else if (dir == "" && fileName != "")
                {
                    // Archivos en el directorio raíz
                    path = destDir;
                }
                else if (dir != "" && fileName != "")
                {
                    // Archivos en la subcarpeta de primer nivel bajo el directorio raíz
                    if (dir.IndexOf("\\") > 0)
                    {
                        // Especifique la ruta para guardar el archivo
                        path = destDir + "\\" + dir;
                    }
                }
                if (dir == rootDir)
                {
                    // Juzgar si es un archivo que debe guardarse en el directorio raíz
                    path = destDir + "\\" + rootDir;
                }

                // Los siguientes son los pasos básicos para descomprimir el archivo zip
                // Idea básica: recorra todos los archivos del archivo comprimido y cree el mismo archivo
                if (fileName != String.Empty)
                {
                    FileStream fs = File.Create(path + "\\" + fileName);
                    int size = 2048;
                    byte[] data = new byte[2048];
                    while (true)
                    {
                        size = inputstream.Read(data, 0, data.Length);
                        if (size > 0)
                        {
                            fs.Write(data, 0, size);
                        }
                        else
                        {
                            break;
                        }
                    }
                    fs.Close();
                }
            }
            inputstream.Close();
        }







        #endregion

        private void button2_Click(object sender, EventArgs e)
        {

            //[*COMPRIME ARCHIVOS QUE SE ENCUENTREN EN LA RUTA*]
            string zipFilePath = @"C:\Users\iruiz\Documents";
            DirectoryInfo ZipdirInfo = new DirectoryInfo(zipFilePath);
           
           foreach (FileInfo zipFilesInfo in ZipdirInfo.GetFiles())
                 {

                     ZipFiles(zipFilesInfo);

                 }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //string ftpURI = "";
            string ftpServerIP = "ftp.bluewebsoft.com";
            string ftpUserID = "bnexttranbox";
            string ftpPassword = "Tr4nBnxt";
            string ftpPort = "21";

            var ftpURI = "ftp://ftp.bluewebsoft.com:21/";

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpURI);
            request.Method = WebRequestMethods.Ftp.DownloadFile;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential("bnexttranbox", "Tr4nBnxt");

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();


            long contentLength = response.ContentLength;


            /*Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            Console.WriteLine(reader.ReadToEnd());

            Console.WriteLine($"Download Complete, status {response.StatusDescription}");

            reader.Close();*/

            /* using (MemoryStream stream=new MemoryStream()) {
                 response.GetResponseStream().CopyTo(stream);
                 response.AddHeader("content-disposition","attachment;filename");
                 response.IsFromCache.SetCacheability(HttpCacheability.Nocache);
                 response.BinaryWrite(stream.ToArray());
                 response.end();

             }*/

          //response.end();
          response.Close();
            //}
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Get the object used to communicate with the server.
            //FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://ftp.bluewebsoft.com:21/");
            FtpWebRequest dirFtp = ((FtpWebRequest)FtpWebRequest.Create("ftp://ftp.bluewebsoft.com:21/"));

            // Los datos del usuario (credenciales)
            NetworkCredential cr = new NetworkCredential("bnexttranbox", "Tr4nBnxt");
            dirFtp.Credentials = cr;

            // El comando a ejecutar usando la enumeración de WebRequestMethods.Ftp
            dirFtp.Method = WebRequestMethods.Ftp.DownloadFile;

            // Obtener el resultado del comando
            StreamReader reader =
                new StreamReader(dirFtp.GetResponse().GetResponseStream());

            // Leer el stream
            string res = reader.ReadToEnd();

            // Mostrarlo.
            //Console.WriteLine(res);

            // Guardarlo localmente con la extensión .txt
            string ficLocal = Path.Combine(@"c:\", Path.GetFileName("ftp://ftp.bluewebsoft.com:21/"+ "BNEXT864200820.txt.gz"));
            StreamWriter sw = new StreamWriter(ficLocal, false, Encoding.UTF8);
            sw.Write(res);
            sw.Close();

            // Cerrar el stream abierto.
            reader.Close();
        }

        private void bton_sftp_Click(object sender, EventArgs e)
        {
                var request = new WebClient { Credentials = new NetworkCredential("demo", "password") };

               // var result = new List<ExternalServiceFtpDto>();

                var date = DateTime.Now;
                var ftpAddressFile = string.Concat("/put/example", "/", "readme.txt");

                byte[] newFileData;
                
               var result = new List<string>();

            using (var sftp = new SftpClient("test.rebex.net", 22, "demo", "password"))
                {
                    sftp.Connect();

                    try
                    {
                    /*  newFileData = sftp.ReadAllBytes(ftpAddressFile);
                      var fileString = Encoding.UTF8.GetString(newFileData);*/


                  //  var files = sftp.ListDirectory("/" + "put/example");
                    var files = sftp.ListDirectory("/pub/example");

                    sftp.DeleteFile("/readme.txt");


                    foreach (var file in files)
                    {
                        if (!string.IsNullOrWhiteSpace(file.Name) && file.Name.EndsWith(".txt", StringComparison.InvariantCultureIgnoreCase))
                        {
                            result.Add(file.Name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Last());
                        }
                    }

                }
                    catch (Exception exception)
                    {
                     //   return result;
                    }

                    sftp.Disconnect();

               
            }

            request.Dispose();

           // return result;

        }

        private void btSFTPbws_Click(object sender, EventArgs e)
        {
          //  var request = new WebClient { Credentials = new NetworkCredential("bnextpv", "7n4-3U20SXt9Aew") };

            // var result = new List<ExternalServiceFtpDto>();

            var date = DateTime.Now;
           // var ftpAddressFile = string.Concat("/Out/TEST", "/", "ArchivoTEST_borrar.txt");

            byte[] newFileData;

            var result = new List<string>();

            using (var sftp = new SftpClient("sftp.bluewebsoft.com", 22, "bnextpv", "7n4-3U20SXt9Aew"))
            {
                sftp.Connect();

                try
                {
                    /*  newFileData = sftp.ReadAllBytes(ftpAddressFile);
                      var fileString = Encoding.UTF8.GetString(newFileData);*/


                    //  var files = sftp.ListDirectory("/" + "put/example");
                    var files = sftp.ListDirectory("/Out/TEST/");

                        //borrar archivo de pruebas en sftp 


                    foreach (var file in files)
                    {
                        if (!string.IsNullOrWhiteSpace(file.Name) && file.Name.EndsWith(".txt", StringComparison.InvariantCultureIgnoreCase))
                        {
                            result.Add(file.Name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Last());

                            sftp.DeleteFile("/Out/TEST/"+ file.Name);
                        }
                    }

                }
                catch (Exception exception)
                {
                    //   return result;
                }

                sftp.Disconnect();


            }

          //  request.Dispose();

            // return result;
        }
    }
}
