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
            string fileId = "";
            do
            {
                fileId = ListFiles(service, ref pageToken);
            } while (pageToken != null);

            SpreadsheetService = sheetService;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(fileId));
        
            Console.Read();
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

                    if (file.Name.Equals("Ekonomi_2020"))
                    {
                        foundIndex = index;
                        fileId = file.Id;
                    }
                    index++;
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            files.ElementAt(foundIndex);

            string currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return fileId;
        }

    }
}
