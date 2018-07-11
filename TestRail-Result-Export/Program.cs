using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Data;
using System.Threading;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;

namespace TestRailResultExport
{
	public class MainClass
	{
		public static string milestoneID = "";
		public static List<string> suiteIDs = new List<string>();
        public static List<string> suiteNames = new List<string>();
        public static List<string> suiteInPlanIDs = new List<string>();
		public static List<string> runIDs = new List<string>();
        public static List<string> runNames = new List<string>();
        public static List<Run> runs = new List<Run>();
        public static List<int> allCaseIDs = new List<int>();
        public static List<int> caseIDsInMilestone = new List<int>(); //case IDs that have been run
        public static List<string> configs = new List<string>();
        public static List<Run> runConfigs = new List<Run>();

        public static int numberPassed;
        public static int numberFailed;
        public static int numberBlocked;

		private static readonly IConfigReader _configReader = new ConfigReader();

        public struct Test
        {
            public string SuiteID;
            public string SuiteName;
            public int RunID;
            public string RunName;
            public string isRunCompleted;
            public string TestID;
            public int CaseID;
            public string Title;
            public string Status;
            public string Defects;
            public string Comment;
            public string Config;
            public string EditorVersion;
			public double elapsedTimeInSeconds;
            public string identifier;
        }

        public struct Case
        {
            public string SuiteID;
            public string SuiteName;
            public string CreatedOn;
            public string UpdatedOn;
            public string Section;
            public int CaseID;
            public string CaseName;
            public string Status;
            public string Type;
            public string TemplateStatus;
        }

        public struct Suite
        {
            public string SuiteID;
            public string SuiteName;
        }

        public struct Run
        {
            public string RunID;
            public string Config;
            public string isCompleted;
        }

		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			APIClient client = ConnectToTestrail();

            //SheetsService sheetsService = GoogleSheets.ConnectToGoogleSheets();
            //DriveService service = GoogleDrive.ConnectToGoogleDrive();

            //GetAllTests(client, 3, "130");
            GetAllTests(client, 3, args[0], args[1]); //Milestone ID and project ID must be entered as cmd line arg

            //GoogleDrive.UploadCsvAsSpreadsheet(service);
            //GoogleDrive.CopyToSheet(sheetsService);

        }

		private static APIClient ConnectToTestrail()
		{
			APIClient client = new APIClient("https://qatestrail.hq.unity3d.com");
			client.User = _configReader.TestRailUser;
			client.Password = _configReader.TestRailPass;
			return client;
		}

        /// <summary>
        /// Retrieves TestRail tests (both in and out of plans)
        /// </summary>
        /// <param name="previousResults">Number of previous results to include.</param>
		private static void GetAllTests(APIClient client, int previousResults, string milestoneID, string projectID)
		{
            //Console.WriteLine("Enter milestone ID: ");
            //milestoneID = Console.ReadLine();
            Console.WriteLine("Getting data from TestRail");
            Console.WriteLine("Milestone ID: " + milestoneID);

            JArray c = AccessTestRail.GetRunsForMilestone(client, milestoneID);
            JArray planArray = AccessTestRail.GetPlansForMilestone(client, milestoneID);
            //The response includes an array of test plans. Each test plan in this list follows the same format as get_plan, except for the entries field which is not included in the response.

            JArray caseTypes = AccessTestRail.GetCaseTypes(client); // This JArray will be used when evaluating case types
            JArray statusArray = AccessTestRail.GetStatuses(client); // This will be used for evaluating status IDs

            List<string> runInPlanIds = AccessTestRail.GetRunsInPlan(planArray, client, suiteInPlanIDs, runNames, runs);

            List<Case> listOfCases = new List<Case>();
            List<Test> listOfTests = new List<Test>();
            List<Suite> listOfSuites = new List<Suite>();

            JArray suitesArray = AccessTestRail.GetSuitesInProject(client, projectID);

            for (int i = 0; i < suitesArray.Count; i++)
            {
                JObject arrayObject = suitesArray[i].ToObject<JObject>();
                string id = arrayObject.Property("id").Value.ToString();
                string suiteName = arrayObject.Property("name").Value.ToString(); //create list of suiteNames to use later

                Suite newSuite;
                newSuite.SuiteID = id;
                newSuite.SuiteName = suiteName;
                listOfSuites.Add(newSuite);


                JArray casesArray = AccessTestRail.GetCasesInSuite(client, projectID, id);
                listOfCases = CreateListOfCases(client, caseTypes, casesArray, listOfCases, id, suiteName);
            }


            AccessTestRail.GetSuitesAndRuns(c, suiteIDs, runIDs, runNames, runs);

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
                ostrm = new FileStream("Tests" + milestoneID + ".csv", FileMode.OpenOrCreate, FileAccess.Write);
				writer = new StreamWriter(ostrm);
			}
			catch (Exception e)
			{
				Console.WriteLine("Cannot open Tests.csv for writing");
				Console.WriteLine(e.Message);
				return;
			}
			Console.SetOut(writer);

			for (int i = 0; i < runIDs.Count; i++)
			{
                JArray testsArray = AccessTestRail.GetTestsInRun(client, runIDs[i]);

                string testID = "";
                int caseID = 0;
                string title = "";
                string status = "";

                for (int j = 0; j < testsArray.Count; j++)
                {
                    JObject testObject = testsArray[j].ToObject<JObject>();
                    testID = testObject.Property("id").Value.ToString();

                    if (testObject.Property("case_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("case_id").Value.ToString()))
                    {
                        caseID = Int32.Parse(testObject.Property("case_id").Value.ToString());
                    }

                    if (!caseIDsInMilestone.Contains(caseID))
                    {
                        caseIDsInMilestone.Add(caseID);
                    }

                    title = testObject.Property("title").Value.ToString();
                    status = StringManipulation.GetStatus(statusArray, testObject.Property("status_id").Value.ToString());

					string suiteName = "";

                    // Some suites have been deleted, but the tests and runs remain
					if (suiteIDs[i] != "0")
					{
                        Suite currentSuite = listOfSuites.Find(x => x.SuiteID == suiteIDs[i]);
                        // Get the suite_id that corresponds to the run_id
                        suiteName = currentSuite.SuiteName;
					}
					else
					{
						suiteName = "deleted";
					}

                    // Get the most recent defects/bugs and comments on the test
                    string defects = "";
                    string comment = "";
                    string editorVersion = "";
					string elapsedTime = "";

                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
						elapsedTime = resultObject.Property("elapsed").Value.ToString();
                    }

                    if (comment.Length > 99)
                    {
                        comment = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(comment.Substring(0, 100)));
                    }

                    if (comment.Contains('"'))
                    {
                        comment = comment.Replace('"', ' ');
                    }
                    else if (comment.Contains(','))
                    {
                        comment = comment.Replace(',', ' ');
                    }
                    else if (comment.Contains('\n'))
                    {
                        comment = comment.Replace('\n', ' ');
                    }

                    Run currentRun = runs.Find(x => x.RunID == runIDs[i]);

                    int elapsedTimeInSeconds = 0;
                    if (elapsedTime != "")
                    {
                        elapsedTimeInSeconds = StringManipulation.ConvertTimespanStringToSeconds(elapsedTime);
                    }


                    Test currentTest;
                    currentTest.SuiteID = suiteIDs[i];
                    currentTest.SuiteName = suiteName;
                    currentTest.RunID = Int32.Parse(runIDs[i]);
                    currentTest.RunName = runNames[i]; // use currentRun
                    currentTest.isRunCompleted = currentRun.isCompleted;
                    currentTest.TestID = testID;
                    currentTest.CaseID = caseID;
                    currentTest.Title = title;
                    currentTest.Status = status;
                    currentTest.Defects = defects;
                    currentTest.Comment = comment;
                    currentTest.Config = ""; // Configs don't exist for runs outside of plans!!!!
                    currentTest.EditorVersion = StringManipulation.GetEditorVersion(editorVersion);
					currentTest.elapsedTimeInSeconds = elapsedTimeInSeconds;
                    currentTest.identifier = caseID + "_" + testID;

					listOfTests.Add(currentTest);
                }

			}

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
                JArray testsArray = AccessTestRail.GetTestsInRun(client, runInPlanIds[i]);

				string testID = "";
				int caseID = 0;
				string title = "";
				string status = "";

				for (int j = 0; j < testsArray.Count; j++)
				{
					JObject testObject = testsArray[j].ToObject<JObject>();

                    testID = testObject.Property("id").Value.ToString();

                    if (testObject.Property("case_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("case_id").Value.ToString()))
                    {
                        caseID = Int32.Parse(testObject.Property("case_id").Value.ToString());
                    }
                    caseIDsInMilestone.Add(caseID);

					title = testObject.Property("title").Value.ToString();
                    status = StringManipulation.GetStatus(statusArray, testObject.Property("status_id").Value.ToString());

					if (status == "Passed")
					{
						numberPassed++;
					}
					else if (status == "Failed")
					{
						numberFailed++;
					}
					else if (status == "Blocked")
					{
						numberBlocked++;
					}

					string suiteName = "";

					// Some suites have been deleted, but the tests and runs remain
					if (suiteInPlanIDs[i] != "0")
					{
                        Suite currentSuite = listOfSuites.Find(x => x.SuiteID == suiteInPlanIDs[i]);
                        suiteName = currentSuite.SuiteName;
					}
					else
					{
						suiteName = "deleted";
					}

                    // Get the most recent defects/bugs and comments on the test
                    string defects = "";
                    string comment = "";
                    string editorVersion = "";
					string elapsedTime = "";

                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
						elapsedTime = resultObject.Property("elapsed").Value.ToString();
                    }

                    // Find config for runID
                    Run currentRun = runs.Find(o => o.RunID == runInPlanIds[i]);
                    string config = currentRun.Config;
                    string runID = currentRun.RunID;

                    
                    if (config.Contains('"'))
                    {
                        config = config.Replace('"', ' ');
                    }

                    if (config.Contains(','))
                    {
                        int index = config.IndexOf(',');
                        config = config.Substring(0, index);
                    }

                    currentRun.Config = config;

                    configs.Add(config);
                    runConfigs.Add(currentRun);

                    if (comment.Length > 99)
                    {
                        comment = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(comment.Substring(0, 100)));
                    }

                    if (comment.Contains('"'))
                    {
                        comment = comment.Replace('"', ' '); 
                    }
                    else if (comment.Contains(','))
                    {
                        comment = comment.Replace(',', ' ');
                    }
                    else if (comment.Contains(Environment.NewLine))
                    {
                        comment = comment.Replace(Environment.NewLine, " ");
                    }

                    int elapsedTimeInSeconds = 0;
                    if (elapsedTime != "")
                    {
                        elapsedTimeInSeconds = StringManipulation.ConvertTimespanStringToSeconds(elapsedTime);
                    }

                    Test currentTest;
                    currentTest.SuiteID = suiteInPlanIDs[i];
                    currentTest.SuiteName = suiteName;
                    currentTest.RunID = Int32.Parse(runInPlanIds[i]);
                    currentTest.RunName = runNames[i];
                    currentTest.isRunCompleted = currentRun.isCompleted;
                    currentTest.TestID = testID;
                    currentTest.CaseID = caseID;
                    currentTest.Title = title;
                    currentTest.Status = status;
                    currentTest.Defects = defects;
                    currentTest.Comment = comment;
                    currentTest.Config = config;
                    currentTest.EditorVersion = StringManipulation.GetEditorVersion(editorVersion);
					currentTest.elapsedTimeInSeconds = elapsedTimeInSeconds;
                    currentTest.identifier = caseID + "_" + testID;

                    listOfTests.Add(currentTest);
				}

			}
            List<Test> sortedList = SortListOfTests(listOfTests);

            string csvOfTests = CreateCSVOfTests(sortedList, previousResults, listOfCases);
            Console.WriteLine(csvOfTests);

            //GoogleSheets.OutputTestsToGoogleSheets(sortedList, previousResults, listOfCases);

            Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}

        public static List<Case> CreateListOfCases(APIClient client, JArray caseTypes, JArray casesArray, List<Case> listOfCases, string suiteID, string suiteName)
        {
            for (int i = 0; i < casesArray.Count; i++)
            {
                JObject arrayObject = casesArray[i].ToObject<JObject>();

                allCaseIDs.Add(Int32.Parse(arrayObject.Property("id").Value.ToString()));

                string caseID = arrayObject.Property("id").Value.ToString();
                string caseName = arrayObject.Property("title").Value.ToString();
                string caseType = arrayObject.Property("type_id").Value.ToString();
                //string templateStatus = arrayObject.Property("custom_case_status").Value.ToString();
                string sectionID = arrayObject.Property("section_id").Value.ToString();

                string createdOn = arrayObject.Property("created_on").Value.ToString();
                string updatedOn = arrayObject.Property("updated_on").Value.ToString();

                JObject section = AccessTestRail.GetSection(client, sectionID);
                string sectionName = AccessTestRail.GetSectionName(section);

                if (sectionName.Contains(","))
                {
                    StringBuilder sb = new StringBuilder(sectionName);
                    sb[sectionName.IndexOf(',')] = ' ';
                    sectionName = sb.ToString();
                }

                Case newCase;
                newCase.SuiteID = suiteID;
                newCase.SuiteName = suiteName;
                //newCase.CreatedOn = DateTimeOffset.FromUnixTimeSeconds(long.Parse(createdOn)).Date.ToString();
                //newCase.UpdatedOn = DateTimeOffset.FromUnixTimeSeconds(long.Parse(updatedOn)).Date.ToString();
                newCase.CreatedOn = createdOn;
                newCase.UpdatedOn = updatedOn;
                newCase.Section = sectionName;
                newCase.CaseID = Int32.Parse(caseID);
                newCase.CaseName = caseName;
                newCase.Status = StringManipulation.IsInvalid(arrayObject);
                newCase.Type = StringManipulation.GetCaseType(caseTypes, caseType);
                newCase.TemplateStatus = "";

                listOfCases.Add(newCase);
            }
            return listOfCases;
        }

        public static string CreateCSVOfTests(List<Test> sortedList, int previousResults, List<Case> listOfCases)
		{
            //Console.WriteLine("Creating CSV");
            StringBuilder csv = new StringBuilder();
			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}", "Suite Name", "Case ID", "Run Name", "Run ID", "Complete", "Identifier", "Section", "Title", "Created On", "Updated On", "Config", "Case Type", "Editor Version", "Last Defects", "Last Comment", "Last Run Result", "Elapsed Time", "\n");
			csv.Append(header);
            List<int> passValues = new List<int>();
            for (int i = 0; i < sortedList.Count; i++)
            {
                Test testObject = sortedList[i];
                Case caseObject = listOfCases.Find(x => x.CaseID == testObject.CaseID); //finding the case that matches the test


                if (testObject.CaseID != 0)
                {
                    if (i != 0)
                    {
                        csv.Append("\n"); //removes the blank row between the headings and the first result
                    }
					string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},", "\"" + testObject.SuiteName + "\"", testObject.CaseID, "\"" + testObject.RunName + "\"", testObject.RunID, testObject.isRunCompleted, testObject.identifier, "\"" + caseObject.Section + "\"", "\"" + testObject.Title + "\"", caseObject.CreatedOn, caseObject.UpdatedOn, "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"", testObject.elapsedTimeInSeconds.ToString());

                    csv.Append(line);
                }

			}
            csv.Append("\n");
            csv.Append("\n");

            for (int k = 0; k < allCaseIDs.Count; k++)
            {
                if (!caseIDsInMilestone.Contains(allCaseIDs[k]))
                {
                    List<Case> sortedListOfCases = SortListOfCases(listOfCases);

                    Case caseNotRun = sortedListOfCases.Find(x => x.CaseID == allCaseIDs[k]);

					string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}", caseNotRun.SuiteName, caseNotRun.CaseID, "Not included in test run", "", "false", caseNotRun.CaseID + "_00", "\"" + caseNotRun.Section + "\"", "\"" + caseNotRun.CaseName + "\"", caseNotRun.CreatedOn, caseNotRun.UpdatedOn, "", "\"" + caseNotRun.Type + "\"", "Never Tested", "", "", "Untested", "0", "\n");
                    csv.Append(line);
                }

            }

			return csv.ToString();
		}

        /// <summary>
        /// Sorts the list of tests by case_id and then run_id
        /// </summary>
        private static List<Test> SortListOfTests(List<Test> listOfTests)
        {
            List<Test> sortedList = listOfTests.OrderBy(o => o.SuiteName).ThenBy(o => o.CaseID).ThenBy(o => o.Config).ToList();
            return sortedList;
        }

        private static List<Case> SortListOfCases(List<Case> listOfCases)
        {
            List<Case> sortedList = listOfCases.OrderByDescending(o => o.CaseID).ToList();
            return sortedList;
        }
	}
}
