
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

        public static void AppendToRow(SheetsService service, ValueRange valueRange, List<ValueRange> data, List<object> appendedRowInGoogleSheet, string sheetID)
        {
            valueRange.Values = new List<IList<object>> { appendedRowInGoogleSheet };
            data.Add(valueRange);

            Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();

            requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
            requestBodyUpdate.Data = data;

            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, sheetID);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            BatchUpdateValuesResponse responseUpdate = request.Execute();
        }

        //public static void CreateNewRow(SheetsService service, ValueRange valueRange, List<ValueRange> data, List<object> rowInGoogleSheet, string sheetID)
        //{
        //    valueRange.Values = new List<IList<object>> { rowInGoogleSheet };
        //    data.Add(valueRange);

        //    Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();
        //    requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
        //    requestBodyUpdate.Data = data;

        //    SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, sheetID);

        //    Data.BatchUpdateValuesResponse responseUpdate = request.Execute();
        //}

        public static void OutputTestsToGoogleSheets(List<MainClass.Test> sortedList, int previousResults, List<MainClass.Case> listOfCases)
		{
            SheetsService service = ConnectToGoogleSheets();
			// The new values to apply to the spreadsheet.
			List<ValueRange> data = new List<ValueRange>();

            int rowNum = 1;

			ValueRange valueRange = new ValueRange();

			int count = 0;
            List<int> passValues = new List<int>();

			for (int i = 0; i < sortedList.Count; i++)
			{
                rowNum++;
                string range = "A" + rowNum + ":Z" + rowNum;
                valueRange.Range = range;

                //MainClass.Test arrayObject = sortedList[i];

                MainClass.Test testObject = sortedList[i];
                MainClass.Case caseObject = listOfCases.Find(x => x.CaseID == testObject.CaseID);

                if (testObject.CaseID != 0)
                {
                    if (i != 0)
                    {
                        // check if the case_id is the same as the one above it
                        if (testObject.CaseID == sortedList[i - 1].CaseID && testObject.Config == sortedList[i - 1].Config)
                        {
                            count++;
                            if (count < previousResults)
                            {

                                string passRate = "";
                                //string line = string.Format("{0},", testObject.Status);
                                List<object> line = new List<object>() { testObject.Status };
                                // 2) add the status to the same list
                                if (testObject.Status == "Passed")
                                {
                                    passValues.Add(100);
                                }
                                else
                                {
                                    passValues.Add(0);
                                }

                                //csv.Append(line);
                                AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                                // if (count-1)=previousResults, calculate pass rate using the small list of pass values
                                if (count == (previousResults - 1))
                                {
                                    // eg sum(passvalues) / previousResults
                                    int sumOfValues = passValues.Sum();
                                    passRate = (sumOfValues / previousResults).ToString();
                                    //csv.Append(string.Format("{0},", passRate + "%"));
                                    List<object> appendedRowInGoogleSheet = new List<object>() { passRate + "%" };

                                    AppendToRow(service, valueRange, data, appendedRowInGoogleSheet, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                                }
                            }
                        }
                        else
                        {
                            // Some values get reset here because this is a brand new line and a new case
                            passValues.Clear();
                            count = 0;
                            if (i != 0)
                            {
                                List<object> newLine = new List<object>() { "\n" };
                                //csv.Append("\n"); 
                                //removes the blank row between the headings and the first result
                                AppendToRow(service, valueRange, data, newLine, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                            }

                            //string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                            List<object> line = new List<object>() { testObject.SuiteName, testObject.Title, testObject.Config, caseObject.Type, StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, testObject.Defects, testObject.Comment, testObject.Status };



                            // if its a pass, value is 100
                            if (testObject.Status == "Passed")
                            {
                                passValues.Add(100);
                            }
                            else
                            {
                                passValues.Add(0);
                            }
                            //csv.Append(line);
                            AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                        }
                    }
                    else
                    {
                        // Some values get reset here because this is a brand new line and a new case
                        passValues.Clear();
                        count = 0;
                        if (i != 0)
                        {
                            List<object> newLine = new List<object>() { "\n" };
                            //csv.Append("\n"); 
                            //removes the blank row between the headings and the first result
                            AppendToRow(service, valueRange, data, newLine, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                        }

                        //string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                        List<object> line = new List<object>() { testObject.SuiteName, testObject.Title, testObject.Config, caseObject.Type, StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, testObject.Defects, testObject.Comment, testObject.Status };



                        // if its a pass, value is 100
                        if (testObject.Status == "Passed")
                        {
                            passValues.Add(100);
                        }
                        else
                        {
                            passValues.Add(0);
                        }
                        //csv.Append(line);
                        AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                    }
                }
			}

			//return csv.ToString();
		}


        public static void OutputTestsToGoogleSheets_Light(List<MainClass.Test> sortedList, int previousResults)
        {
            SheetsService service = ConnectToGoogleSheets();
            // The new values to apply to the spreadsheet.
            List<ValueRange> data = new List<ValueRange>();

            int rowNum = 1;

            ValueRange valueRange = new ValueRange();

            int count = 0;
            List<int> passValues = new List<int>();

            for (int i = 0; i < sortedList.Count; i++)
            {
                rowNum++;
                string range = "A" + rowNum + ":Z" + rowNum;
                valueRange.Range = range;

                //MainClass.Test arrayObject = sortedList[i];

                MainClass.Test testObject = sortedList[i];
                //MainClass.Case caseObject = listOfCases.Find(x => x.CaseID == testObject.CaseID);

                if (testObject.CaseID != 0)
                {
                    if (i != 0)
                    {
                        // check if the case_id is the same as the one above it
                        if (testObject.CaseID == sortedList[i - 1].CaseID)
                        {
                            count++;
                            if (count < previousResults)
                            {

                                string passRate = "";
                                //string line = string.Format("{0},", testObject.Status);
                                List<object> line = new List<object>() { testObject.Status };
                                // 2) add the status to the same list
                                if (testObject.Status == "Passed")
                                {
                                    passValues.Add(100);
                                }
                                else
                                {
                                    passValues.Add(0);
                                }

                                //csv.Append(line);
                                AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                                // if (count-1)=previousResults, calculate pass rate using the small list of pass values
                                if (count == (previousResults - 1))
                                {
                                    // eg sum(passvalues) / previousResults
                                    int sumOfValues = passValues.Sum();
                                    passRate = (sumOfValues / previousResults).ToString();
                                    //csv.Append(string.Format("{0},", passRate + "%"));
                                    List<object> appendedRowInGoogleSheet = new List<object>() { passRate + "%" };

                                    AppendToRow(service, valueRange, data, appendedRowInGoogleSheet, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                                }
                            }
                        }
                        else
                        {
                            // Some values get reset here because this is a brand new line and a new case
                            passValues.Clear();
                            count = 0;
                            if (i != 0)
                            {
                                List<object> newLine = new List<object>() { "\n" };
                                //csv.Append("\n"); 
                                //removes the blank row between the headings and the first result
                                AppendToRow(service, valueRange, data, newLine, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                            }

                            //string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                            List<object> line = new List<object>() { testObject.SuiteName, testObject.Title, testObject.Status };



                            // if its a pass, value is 100
                            if (testObject.Status == "Passed")
                            {
                                passValues.Add(100);
                            }
                            else
                            {
                                passValues.Add(0);
                            }
                            //csv.Append(line);
                            AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                        }
                    }
                    else
                    {
                        // Some values get reset here because this is a brand new line and a new case
                        passValues.Clear();
                        count = 0;
                        if (i != 0)
                        {
                            List<object> newLine = new List<object>() { "\n" };
                            //csv.Append("\n"); 
                            //removes the blank row between the headings and the first result
                            AppendToRow(service, valueRange, data, newLine, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                        }

                        //string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", StringManipulation.GetTemplateStatus(caseObject.TemplateStatus), testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                        List<object> line = new List<object>() { testObject.SuiteName, testObject.Title, testObject.Status };



                        // if its a pass, value is 100
                        if (testObject.Status == "Passed")
                        {
                            passValues.Add(100);
                        }
                        else
                        {
                            passValues.Add(0);
                        }
                        //csv.Append(line);
                        AppendToRow(service, valueRange, data, line, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");
                    }
                }
            }

            //return csv.ToString();
        }

        public static void UploadCsvToSheet(SheetsService service, string csv, string spreadsheetId)
        {
            //SheetsService service = ConnectToGoogleSheets();
            int rowNum = 1;
            ValueRange valueRange = new ValueRange();

            string range = "A" + rowNum + ":Z" + rowNum;
            valueRange.Range = range;


            PasteDataRequest pasteDataRequest = new PasteDataRequest();
            pasteDataRequest.Data = csv;
            pasteDataRequest.Coordinate = new GridCoordinate();
            //pasteDataRequest.Type = pas


            Request request = new Request();
            request.PasteData = pasteDataRequest;

            //SpreadsheetsResource.CreateRequest req = service.Spreadsheets.Values.BatchUpdate(pasteDataRequest, spreadsheetId);

            //SpreadsheetsResource.BatchUpdateRequest req = service.Spreadsheets.BatchUpdate(request.)
            
        }

        public static void WriteToSheet(List<object> oblist)
		{
            SheetsService service = ConnectToGoogleSheets();

			// The new values to apply to the spreadsheet.
			List<ValueRange> data = new List<ValueRange>();

			string range = "A:Z";
			ValueRange valueRange = new ValueRange();
			valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS
			valueRange.Range = range;

            //List<object> oblist = new List<object>(); //treat each row of the csv as an object?
			valueRange.Values = new List<IList<object>> { oblist };

			data.Add(valueRange);

			Data.BatchUpdateValuesRequest requestBodyUpdate = new Data.BatchUpdateValuesRequest();
			requestBodyUpdate.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString();
			requestBodyUpdate.Data = data;

            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBodyUpdate, "1to1HtFx5WoAjU07OLWMa1gqm8WjGG9T3PFK-7cDEjKA");

			// To execute asynchronously in an async method, replace `request.Execute()` as shown:
			Data.BatchUpdateValuesResponse responseUpdate = request.Execute();

			//Console.WriteLine(spreadsheetURL);
		}
	}
}