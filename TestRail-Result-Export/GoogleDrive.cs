using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Net;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Google.Apis.Sheets.v4;

namespace TestRailResultExport
{
    public class GoogleDrive
    {
        static string ApplicationName = "TestRail-Results-Export";
        static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile };

        public static DriveService ConnectToGoogleDrive()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive.googleapis.com-testrail.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //// Define parameters of request.
            //FilesResource.ListRequest listRequest = service.Files.List();
            //listRequest.PageSize = 10;
            //listRequest.Fields = "nextPageToken, files(id, name)";

            //// List files.
            //IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            //Console.WriteLine("Files:");
            //if (files != null && files.Count > 0)
            //{
            //    foreach (var file in files)
            //    {
            //        Console.WriteLine("{0} ({1})", file.Name, file.Id);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No files found.");
            //}
            //Console.Read();

            return service;
        }

        public static void UploadCsvAsSpreadsheet(DriveService driveService)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = "CSV Report",
                MimeType = "application/vnd.google-apps.spreadsheet"
            };
            //FilesResource.CreateMediaUpload request;
            FilesResource.UpdateMediaUpload request;
            using (var stream = new System.IO.FileStream("Tests.csv", System.IO.FileMode.Open))
            {
                //request = driveService.Files.Create(fileMetadata, stream, "text/csv");
                request = driveService.Files.Update(fileMetadata, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA", stream, "text/csv");
                request.Fields = "id";

                request.Upload();
            }
            var file = request.ResponseBody;
            Console.WriteLine("File ID: " + file.Id);
        }

        public static void CopyToSheet(SheetsService sheetsService)
        {
            // The ID of the spreadsheet containing the sheet to copy.
            string spreadsheetId = "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA";

            // The ID (gid) of the sheet to copy.
            int sheetId = 400073617;

            // The ID of the spreadsheet to copy the sheet to.
            string destinationSpreadsheetId = "1y0jHtVpdIF6Mr2vBShdrXtiIBcAEJshKqjUdR-Nantk";

            Google.Apis.Sheets.v4.Data.CopySheetToAnotherSpreadsheetRequest requestBody = new Google.Apis.Sheets.v4.Data.CopySheetToAnotherSpreadsheetRequest();
            requestBody.DestinationSpreadsheetId = destinationSpreadsheetId;
            

            SpreadsheetsResource.SheetsResource.CopyToRequest request = sheetsService.Spreadsheets.Sheets.CopyTo(requestBody, spreadsheetId, sheetId);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Google.Apis.Sheets.v4.Data.SheetProperties response = request.Execute();
            // Data.SheetProperties response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));

            //TODO: rename newly created sheet and delete old one?
        }

    }
}
