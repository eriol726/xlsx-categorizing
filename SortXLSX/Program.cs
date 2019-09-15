using ExcelDataReader;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Google.GData.Client;
using Google.GData.Extensions;
using System.Configuration;

namespace SortXLSX
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.Drive };
        static string ApplicationName = "Drive API .NET Quickstart";
        public static SheetsService SpreadsheetService;
        public static string userName = ConfigurationManager.AppSettings["userName"];

        public static object ExcelLibrary { get; private set; }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            UserCredential credential;

            credential = GetCredentials();
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Create Google Sheets API service.
            var sheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //UploadBasicImage("C:/Users/Erik/Pictures/lost_warning.jpg", service);

            string pageToken = null;

            do
            {
                ListFiles(service, ref pageToken);
            } while (pageToken != null);

            SpreadsheetService = sheetService;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        
            Console.Read();
        }

        private static void UploadBasicImage(string path, DriveService service)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File();
            fileMetadata.Name = Path.GetFileName(path);


            fileMetadata.MimeType = "image/jpeg";
            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream(path, System.IO.FileMode.Open))
            {
                request = service.Files.Create(fileMetadata, stream, "image/jpeg");
                request.Fields = "id";
                request.Upload();
            }

            var file = request.ResponseBody;

            Console.WriteLine("File ID: " + file.Id);

        }

        private static UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    userName,
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        private static string ListFiles(DriveService service, ref string pageToken)
        {
            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            
            Console.WriteLine("Files:");
            int index = 0;
            int foundIndex = 0;
            string fileId = "";
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    
                    if (file.Name.Contains("Ekonomi 2019"))
                    {
                        foundIndex = index;
                        Console.WriteLine("{0}", file.Id);
                        fileId = file.Id;
                    }
                }
                index++;
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            files.ElementAt(foundIndex);

            string currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return fileId;


        }
        private static void DownloadFile(DriveService service, Google.Apis.Drive.v3.Data.File file, string saveTo)
        {
            Console.WriteLine("file: " + file.Id);
            //var request = service.Files.Get(file.Id);
            var stream = new System.IO.MemoryStream();

            Console.WriteLine("file.mimic: " + file.MimeType);
            var request = service.Files.Export(file.Id, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                        {
                            Console.WriteLine(progress.BytesDownloaded);
                            break;
                        }
                    case DownloadStatus.Completed:
                        {
                            Console.WriteLine("Download complete.");
                            SaveStream(stream, saveTo, service, file);
                            break;
                        }
                    case DownloadStatus.Failed:
                        {
                            Console.WriteLine("Download failed.");
                            if (file.ExportLinks.Any())
                                SaveStream(new System.Net.Http.HttpClient().GetStreamAsync(file.ExportLinks.FirstOrDefault().Value).Result, saveTo + @"/" +"file" );
                            
                            break;
                        }
                }
            };
            try
            {
                request.Download(stream);
            }
            catch (Exception ex)

            {
                Console.Write("Error: " + ex.Message);
            }
        }

        private static void SaveStream(Stream result, string v)
        {
            throw new NotImplementedException();
        }

        private static void SaveStream(System.IO.MemoryStream stream, string saveTo, DriveService service, Google.Apis.Drive.v3.Data.File _fileResource)
        {
            System.IO.File.SetAttributes(saveTo, FileAttributes.Normal);

            string fullName = Path.Combine(saveTo, "Ekonomi_2019_downloaded.xlsx");

            using (System.IO.FileStream file = new System.IO.FileStream(fullName, System.IO.FileMode.Create))
            {
                //var x = service.HttpClient.GetByteArrayAsync(_fileResource.);
                //byte[] arrBytes = x.Result;
                //System.IO.File.WriteAllBytes(saveTo, arrBytes);
                stream.WriteTo(file);
            }
        }
        private void BtnOpen_Click(object sender, EventArgs e)
        {
            
        }

    }
}
