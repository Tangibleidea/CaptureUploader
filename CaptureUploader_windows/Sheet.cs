using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CaptureUploader
{


    class Sheet
    {
        String targetURL = "https://docs.google.com/spreadsheets/d/16XngdfajuFrCzi_fIUr6kyNiRRfVM_aVQfIKk38G4Oc/edit#gid=1417473064";
        String targetSheet = "2019.Mar";
        static string[] Scopes = { SheetsService.Scope.Spreadsheets, DriveService.Scope.Drive };
        String ApplicationName = "Capture Uploader";

        String SpreadSheetID = String.Empty;
        Spreadsheet currSheet = null;
//        SheetsService service = null;

        public Sheet()
        {
            OpenSheet();
        }

        private Google.Apis.Sheets.v4.SheetsService GetService_sheetv4()
        {
            //if (service != null)
            //    return;

            UserCredential credential;
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            string clientID = Path.Combine(projectDirectory, "Resources\\credentials.json");

            using (var stream =
                new FileStream(clientID, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/token.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public void OpenSheet()
        {
            var service = GetService_sheetv4();
            Console.WriteLine("Open google sheet file : " + targetURL);

            targetURL = targetURL.Replace("https://", "");
            targetURL = targetURL.Replace("docs.google.com/spreadsheets/d/", "");
            var a = targetURL.Split('/');
            if (a.Length >= 1)
                SpreadSheetID = a[0];

            Console.WriteLine("Target Sheet ID : " + SpreadSheetID);

            // Define request parameters.
            var sheet_metadata = service.Spreadsheets.Get(SpreadSheetID).Execute();
            currSheet = sheet_metadata;
        }

        /// <summary>
        /// 
        /// </summary>
        public void ListNameOfSheets()
        {
            if (currSheet == null)
                return;

            foreach (var S in currSheet.Sheets)
            {
                Console.WriteLine(S.Properties.Title);
            }
        }

        public void GetTheLastSheet()
        {
            var last= currSheet.Sheets.Last();
            
        }

        public void EditSheet(String txt, int row, String col)
        {
            var service = GetService_sheetv4();
            var last = currSheet.Sheets.Last();

            int lastrow= GetLastRow();
            String range = $"{last.Properties.Title}!{col}{row}";

            ValueRange valueRange = new ValueRange();
            var oblist = new List<object>() { txt };
            valueRange.Values = new List<IList<object>> { oblist };
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

            Console.WriteLine("Writing location: " + range);
            Console.WriteLine("Content: " + txt);

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, SpreadSheetID, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
        }

        public int GetLastRow()
        {
            var service = GetService_sheetv4();
            var last = currSheet.Sheets.Last();

            String range = $"{last.Properties.Title}!A1:Z";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadSheetID, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> tableValues = response.Values;
            if (tableValues != null && tableValues.Count > 0)
            {
                //Console.WriteLine("Name, Major");
                int index = 0;
                foreach (var row in tableValues)
                {
                    ++index;
                    if (row.Count == 0)
                        return index;
                }
            }
            return -1;
        }
    }
}
