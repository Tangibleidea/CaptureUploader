using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation; 

namespace CaptureUploader
{
    class Program
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets, DriveService.Scope.Drive };


        static UserCredential credential = null;
        public static Google.Apis.Drive.v3.DriveService GetService_v3()
        {
            if(credential == null)
            {
                string workingDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
                string clientID = Path.Combine(projectDirectory, "Resources\\credentials.json");

                Console.WriteLine("projectDirectory : " + clientID);
                if (!File.Exists(clientID))
                {
                    Console.WriteLine("credentials exists");
                }
                else
                {
                    Console.WriteLine("credentials does not exist");
                }


                string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/token.json");
                
                GoogleCredential GC = new GoogleCredential();
                credential = GC.GetGoogleCredential(Scopes);
            }
            

            //Create Drive API service.
            Google.Apis.Drive.v3.DriveService service = new Google.Apis.Drive.v3.DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GoogleDriveRestAPI-v3",
            });

            return service;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="findTarget"></param>
        /// <returns></returns>
        static private Google.Apis.Drive.v3.Data.File SearchTarget(String findTarget, bool isFolder, String parentID = null)
        {
            var service = GetService_v3();
            string pageToken = null;
            do
            {
                var request = service.Files.List();
                String Query = "name = '" + findTarget;
                if (isFolder)
                    Query += "' and mimeType = 'application/vnd.google-apps.folder'";
                else
                    Query += "' and mimeType != 'application/vnd.google-apps.folder'";

                if (parentID != null)
                    Query += " and '" + parentID + "' in parents";

                Console.WriteLine("Q: " + Query);

                request.Q = Query;
                request.Spaces = "drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;
                var result = request.Execute();
                foreach (var file in result.Files)
                {

                    Console.WriteLine(String.Format(
                            "Path Found: {0} ({1})", file.Name, file.Id));
                    return file;
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            return null;
        }

        
        static private bool CreateGoogleDriveFolder(String name, String parentID)
        {
            try
            {
                var service = GetService_v3();
                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = name,
                    MimeType = "application/vnd.google-apps.folder",
                    Parents = new List<string>
                {
                    parentID
                },
                };
                var request = service.Files.Create(fileMetadata);
                request.Fields = "nextPageToken, files(id, name)";
                var file = request.Execute();

                if (file == null)
                {
                    return false;
                }
                else
                {
                    Console.WriteLine("new Folder ID: " + file.Id);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return true;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arg"></param>
        static private String AccessGoogleDrive(String arg)
        {
            var folderKoreanFilterList = SearchTarget("Korean filter list screenshots", true);
            if (folderKoreanFilterList == null)
            {
                Console.WriteLine("No target to upload.");
                return null;
            }

            DateTime dtToday = DateTime.Today;
            var folderYear = SearchTarget(dtToday.Year.ToString(), true, folderKoreanFilterList.Id);
            if (folderYear == null)
            {
                Console.WriteLine("No target to upload.");

                if (CreateGoogleDriveFolder(dtToday.Year.ToString(), folderKoreanFilterList.Id))
                {
                    folderYear = SearchTarget(dtToday.Year.ToString(), true, folderKoreanFilterList.Id);
                    if(folderYear == null)
                    {
                        Console.WriteLine("No target to upload. retry failed.");
                        return null;
                    }
                }
                else
                {
                    Console.WriteLine("Failed to create a directoy : {0}", dtToday.Year.ToString());
                    return null;

                }
            }

            String name = dtToday.ToString("MMMdd");
            var folderUploadHere = SearchTarget(name, true, folderYear.Id);
            if (folderUploadHere == null)
            {
                if (CreateGoogleDriveFolder(name, folderYear.Id))
                {
                    folderUploadHere = SearchTarget(name, true, folderYear.Id);
                }
            }

            //Console.WriteLine("arg input: " + arg);
            String fileName = Path.GetFileName(arg);
            //Console.WriteLine("fileName input: " + fileName);

            var service = GetService_v3();
            var fileMeta = new Google.Apis.Drive.v3.Data.File()
            {
                Parents = new List<string>
                {
                    folderUploadHere.Id
                },
                Name = fileName
            };

            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream(arg,
                                    System.IO.FileMode.Open))
            {
                request = service.Files.Create(
                    fileMeta, stream, "image/png");
                request.Fields = "id";
                request.Upload();
            }

            var file = request.ResponseBody;
            if(file == null)
            {
                Console.WriteLine("ERR: request.ResponseBody is null.");
                return null;
            }
            Console.WriteLine("File ID: " + file.Id);

            //System.Diagnostics.Process.Start("https://drive.google.com/open?id=" + file.Id);
            return "https://drive.google.com/open?id=" + file.Id.ToString();
        }


        static void Main(string[] args)
        {
            var currURL = string.Empty;
            var task = Task.Run(() =>
            {
                return Chrome.GetCurrentChromeURL();
            });

            bool isCompletedSuccessfully = task.Wait(TimeSpan.FromMilliseconds(5000));
            if (isCompletedSuccessfully)
            {
                currURL= task.Result;
            }
            else
            {
                //throw new TimeoutException("The function has taken longer than the maximum time allowed.");
                Console.WriteLine("The function has taken longer than the maximum time allowed.");
            }
            //var currURL= Chrome.GetCurrentChromeURL();

            string arg1 = String.Empty;
            if (args.Length == 1)
            {
                arg1 = args[0];
            }
            else
            {
                Console.WriteLine("Please input the target image path as an argument: ");
                arg1= Console.ReadLine();
                //throw new ArgumentException("Please input the target image path as an argument.");
                //arg1 = "sample.png";
            }
            Console.WriteLine("Arg: " + arg1);

            String uploadedFileURL = AccessGoogleDrive(arg1);

            if (uploadedFileURL == null)
            {
                Console.ReadKey();
                return;
            }

            Sheet sheet = new Sheet();
            //sheet.ListNameOfSheets();
            int lastrow = sheet.GetLastRow();
            sheet.EditSheet(currURL, lastrow, "B");
            sheet.EditSheet(uploadedFileURL, lastrow, "C");

            //Environment.Exit(0);
        }
    }
}
