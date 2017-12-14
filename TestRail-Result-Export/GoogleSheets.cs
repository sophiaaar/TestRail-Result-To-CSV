
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

using Data = Google.Apis.Sheets.v4.Data;

namespace TestRailResultExport
{
	class GoogleSheets
	{
		static string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static string ApplicationName = "TestRail-Results-Export";

        public static SheetsService ConnectToGoogleSheets()
		{
			UserCredential credential;

			using (var stream =
				new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
			{
				string credPath = System.Environment.GetFolderPath(
					System.Environment.SpecialFolder.Personal);
				credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-testrail.json");

				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.Load(stream).Secrets,
					Scopes,
					"user",
					CancellationToken.None,
					new FileDataStore(credPath, true)).Result;
				Console.WriteLine("Credential file saved to: " + credPath);
			}

			// Create Google Sheets API service.
			var service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

            return service;
		}

        public static void CreateSpreadsheet(SheetsService service)
        {
			Data.Spreadsheet requestBody = new Data.Spreadsheet();

			SpreadsheetsResource.CreateRequest sheetRequest = service.Spreadsheets.Create(requestBody);

			Data.Spreadsheet response = sheetRequest.Execute();
			string spreadsheetID = response.SpreadsheetId;
			string spreadsheetURL = response.SpreadsheetUrl;         
        }

		public static void OutputTestsToGoogleSheets(List<Test> sortedList, int previousResults)
		{
            SheetsService service = ConnectToGoogleSheets();
			// The new values to apply to the spreadsheet.
			List<ValueRange> data = new List<ValueRange>();

            int rowNum = 1;

			ValueRange valueRange = new ValueRange();

			int count = 0;
			for (int i = 0; i < sortedList.Count; i++)
			{
                rowNum++;
                string range = "A" + rowNum + ":Z" + rowNum;
                valueRange.Range = range;

				Test arrayObject = sortedList[i];
				if (i != 0)
				{
					if (arrayObject.CaseID != 0)
					{
						// check if the case_id is the same as the one above it
						if (arrayObject.CaseID == sortedList[i - 1].CaseID)
						{
							count++;
							if (count < previousResults)
							{

								List<object> appendedRowInGoogleSheet = new List<object>() { arrayObject.RunID, arrayObject.TestID, arrayObject.Status };
		                        valueRange.Values = new List<IList<object>> { appendedRowInGoogleSheet };
		                        data.Add(valueRange);

								Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();
								requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
								requestBodyUpdate.Data = data;

								SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, "1U6I_F8TO8AhJsWwy05b1Y1LdUxO7WkLnEPdIC9tHUU0");

								// To execute asynchronously in an async method, replace `request.Execute()` as shown:
								BatchUpdateValuesResponse responseUpdate = request.Execute();

								//List<object> appendedRowInGoogleSheet = new List<object>() { arrayObject.RunID, arrayObject.TestID, arrayObject.Status };
								//valueRange.Values = new List<IList<object>> { appendedRowInGoogleSheet };
								//data.Add(valueRange);

								//AppendCellsRequest requestBodyUpdate = new AppendCellsRequest();
								////requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
								////requestBodyUpdate.Data = data;

								//SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(valueRange, "1U6I_F8TO8AhJsWwy05b1Y1LdUxO7WkLnEPdIC9tHUU0", range);

								//// To execute asynchronously in an async method, replace `request.Execute()` as shown:
								//AppendValuesResponse responseUpdate = request.Execute();
							}
							else
							{
							}
						}
						else
						{
							count = 0;
							//csv.Append("\n");


							//string line = string.Format("{0},{1},{2},{3},{4},{5},{6},", arrayObject.SuiteID, arrayObject.SuiteName, arrayObject.RunID, arrayObject.TestID, arrayObject.CaseID, "\"" + arrayObject.Title + "\"", arrayObject.Status);

							List<object> rowInGoogleSheet = new List<object>() { arrayObject.SuiteID, arrayObject.SuiteName, arrayObject.RunID, arrayObject.TestID, arrayObject.CaseID, "\"" + arrayObject.Title + "\"", arrayObject.Status };
                            valueRange.Values = new List<IList<object>> { rowInGoogleSheet };
                            data.Add(valueRange);

							Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();
							requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
							requestBodyUpdate.Data = data;

							SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, "1U6I_F8TO8AhJsWwy05b1Y1LdUxO7WkLnEPdIC9tHUU0");

							// To execute asynchronously in an async method, replace `request.Execute()` as shown:
							Data.BatchUpdateValuesResponse responseUpdate = request.Execute();

						}
					}
				}
			}

			//return csv.ToString();
		}

        public static void WriteToSheet(SheetsService service, string spreadsheetID, string spreadsheetURL)
		{
			// The new values to apply to the spreadsheet.
			List<ValueRange> data = new List<ValueRange>();

			string range = "A:Z";
			ValueRange valueRange = new ValueRange();
			valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS
			valueRange.Range = range;

			var oblist = new List<object>() { "My Cell Text", "test", "another test" }; //treat each row of the csv as an object?
			valueRange.Values = new List<IList<object>> { oblist };

			data.Add(valueRange);

			Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();
			requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
			requestBodyUpdate.Data = data;

			SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, spreadsheetID);

			// To execute asynchronously in an async method, replace `request.Execute()` as shown:
			Data.BatchUpdateValuesResponse responseUpdate = request.Execute();

			Console.WriteLine(spreadsheetURL);
		}
	}
}