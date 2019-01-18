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
		//public static string milestoneNameGlobal = "";
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
			public string Area;
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
			public double Estimate;
			public double EstimateForecast;
			public double elapsedTimeInSeconds;
			public int CompletedDate;
			public string MilestoneID;
			public string MilestoneName;
            public string identifier;
			//public string isRetest;
        }

        public struct Case
        {
			public string Area;
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
			public double Estimate;
            public double EstimateForecast;
			public string MilestoneName;
			public string UniqueCaseIdentifier; //this is to identify the case when it is moved between projects and the ID changes
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
			public string MilestoneID;
        }

		public struct Milestone
		{
			public string MilestoneID;
			public string MilestoneName;
		}

		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			APIClient client = ConnectToTestrail();

            GetAllTests(client, args[0]); //Milestone ID and project ID must be entered as cmd line arg         
        }

		private static APIClient ConnectToTestrail()
		{
			APIClient client = new APIClient("https://qatestrail.hq.unity3d.com");
			client.User = _configReader.TestRailUser;
			client.Password = _configReader.TestRailPass;
			return client;
		}

		private static void GetAllTests(APIClient client, string projectID)
		{
            //Console.WriteLine("Enter milestone ID: ");
            //milestoneID = Console.ReadLine();
            Console.WriteLine("Getting data from TestRail");
            //Console.WriteLine("Milestone ID: " + milestoneID);
            Console.WriteLine("Project ID: " + projectID);

			//JArray c = AccessTestRail.GetRunsForMilestone(client, projectID, milestoneID);
			//JArray planArray = AccessTestRail.GetPlansForMilestone(client, projectID, milestoneID);
            JArray c = AccessTestRail.GetRuns(client, projectID);
            JArray planArray = AccessTestRail.GetPlans(client, projectID);

            //The response includes an array of test plans. Each test plan in this list follows the same format as get_plan, except for the entries field which is not included in the response.

            JArray caseTypes = AccessTestRail.GetCaseTypes(client); // This JArray will be used when evaluating case types
            JArray statusArray = AccessTestRail.GetStatuses(client); // This will be used for evaluating status IDs

			JArray milestonesArray = AccessTestRail.GetMilestones(client, projectID);

			JObject project = AccessTestRail.GetProject(client, projectID);
			string projectName = project.Property("name").Value.ToString();

			List<Milestone> listOfMilestones = new List<Milestone>();
			for (int k = 0; k < milestonesArray.Count; k++)
            {
                JObject milestoneObject = milestonesArray[k].ToObject<JObject>();

				Milestone milestone;
                milestone.MilestoneID = milestoneObject.Property("id").Value.ToString();
                milestone.MilestoneName = milestoneObject.Property("name").Value.ToString();

				listOfMilestones.Add(milestone);
            }

			//JObject milestoneSingularObject = AccessTestRail.GetMilestone(client, milestoneID);
			//milestoneNameGlobal = milestoneSingularObject.Property("name").Value.ToString();

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

				string milestoneString = "";
                if (projectName.Contains("Unity "))
                {
                    milestoneString = projectName.Remove(0, 6);
                }
                


				listOfCases = CreateListOfCases(client, caseTypes, casesArray, listOfCases, id, suiteName, milestoneString);
            }


            AccessTestRail.GetSuitesAndRuns(c, suiteIDs, runIDs, runNames, runs);

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
				ostrm = new FileStream("Tests" + projectID + ".csv", FileMode.OpenOrCreate, FileAccess.Write);
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
				string area = "";
				string milestoneID = "";
				string estimate = "";
				string estimateForecast = "";

                for (int j = 0; j < testsArray.Count; j++)
                {
                    JObject testObject = testsArray[j].ToObject<JObject>();
                    testID = testObject.Property("id").Value.ToString();

                    if (testObject.Property("area") != null && testObject.Property("area").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("area").Value.ToString()))
                    {
                        area = testObject.Property("area").Value.ToString();
                    }

                    if (testObject.Property("case_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("case_id").Value.ToString()))
                    {
                        caseID = Int32.Parse(testObject.Property("case_id").Value.ToString());
                    }

					if (testObject.Property("milestone_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("milestone_id").Value.ToString()))
                    {
						milestoneID = testObject.Property("milestone_id").Value.ToString();
                    }

					int estimateInSeconds = 0;
					if (testObject.Property("estimate").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("estimate").Value.ToString()))
                    {
                        estimate = testObject.Property("estimate").Value.ToString();
						if (estimate != "")
                        {
							estimateInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimate);
                        }
                    }

					int estimateForecastInSeconds = 0;
                    if (testObject.Property("estimate_forecast").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("estimate_forecast").Value.ToString()))
                    {
						estimateForecast = testObject.Property("estimate_forecast").Value.ToString();
						if (estimateForecast != "")
                        {
							estimateForecastInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimateForecast);
                        }
                    }

                    if (!caseIDsInMilestone.Contains(caseID))
                    {
                        caseIDsInMilestone.Add(caseID);
                    }

                    title = testObject.Property("title").Value.ToString();
                    if (title.Contains(","))
                    {
                        title = title.Replace(',', ' ');
                    }
                    if (title.Contains('"'))
                    {
                        title = title.Replace('"', ' ');
                    }
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
					string completedDate = "";


                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
						elapsedTime = resultObject.Property("elapsed").Value.ToString();
						completedDate = resultObject.Property("created_on").Value.ToString();
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

					string milestoneName = "";
					Milestone currentMilestone = listOfMilestones.Find(m => m.MilestoneID == currentRun.MilestoneID);

					if (!string.IsNullOrWhiteSpace(currentRun.MilestoneID))
                    {
						milestoneName = currentMilestone.MilestoneName;
                    }
                    else
                    {
                        string milestoneString = "";
                        if (projectName.Contains("Unity"))
                        {
                            milestoneString = projectName.Remove(0, 6);
                        }
                        milestoneName = milestoneString;
                    }               

					int completedInt = 0;

                    if (completedDate == "")
                    {
                        completedInt = 0;
                    }
                    else
                    {
                        completedInt = Int32.Parse(completedDate);
                    }


                    Test currentTest;
					currentTest.Area = area;
                    currentTest.SuiteID = suiteIDs[i];
                    currentTest.SuiteName = suiteName;
                    currentTest.RunID = Int32.Parse(runIDs[i]);
                    currentTest.RunName = runNames[i];
                    currentTest.isRunCompleted = currentRun.isCompleted;
                    currentTest.TestID = testID;
                    currentTest.CaseID = caseID;
                    currentTest.Title = title;
                    currentTest.Status = status;
                    currentTest.Defects = defects;
                    currentTest.Comment = comment;
                    currentTest.Config = ""; // Configs don't exist for runs outside of plans!!!!
					currentTest.EditorVersion = AccessTestRail.GetEditorVersion(client, projectID, editorVersion);
					currentTest.Estimate = estimateInSeconds;
					currentTest.EstimateForecast = estimateForecastInSeconds;
					currentTest.elapsedTimeInSeconds = elapsedTimeInSeconds;
					currentTest.CompletedDate = completedInt;
					currentTest.MilestoneID = currentRun.MilestoneID;
					currentTest.MilestoneName = milestoneName;
                    currentTest.identifier = caseID + "_" + testID;
					//currentTest.isRetest = isRetest.ToString();

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
				string area = "";
				string milestoneID = "";
				string estimate = "";
                string estimateForecast = "";

				for (int j = 0; j < testsArray.Count; j++)
				{
					JObject testObject = testsArray[j].ToObject<JObject>();

                    testID = testObject.Property("id").Value.ToString();
                    if (testObject.Property("area") != null && testObject.Property("area").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("area").Value.ToString()))
                    {
                        area = testObject.Property("area").Value.ToString();
                    }

                    if (testObject.Property("case_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("case_id").Value.ToString()))
                    {
                        caseID = Int32.Parse(testObject.Property("case_id").Value.ToString());
                    }

					if (testObject.Property("milestone_id").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("milestone_id").Value.ToString()))
                    {
                        milestoneID = testObject.Property("milestone_id").Value.ToString();
                    }

					int estimateInSeconds = 0;
                    if (testObject.Property("estimate").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("estimate").Value.ToString()))
                    {
                        estimate = testObject.Property("estimate").Value.ToString();
                        if (estimate != "")
                        {
                            estimateInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimate);
                        }
                    }

                    int estimateForecastInSeconds = 0;
                    if (testObject.Property("estimate_forecast").Value != null && !string.IsNullOrWhiteSpace(testObject.Property("estimate_forecast").Value.ToString()))
                    {
                        estimateForecast = testObject.Property("estimate_forecast").Value.ToString();
                        if (estimateForecast != "")
                        {
                            estimateForecastInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimateForecast);
                        }
                    }

                    caseIDsInMilestone.Add(caseID);
                    title = testObject.Property("title").Value.ToString();
                    if (title.Contains(","))
                    {
                        title = title.Replace(',', ' ');
                    }
                    if (title.Contains('"'))
                    {
                        title = title.Replace('"', ' ');
                    }
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
					string completedDate = "";

                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
						elapsedTime = resultObject.Property("elapsed").Value.ToString();
						completedDate = resultObject.Property("created_on").Value.ToString();
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

                    if (comment.Contains(','))
                    {
                        comment = comment.Replace(',', ' ');
                    }

                    if (comment.Contains(Environment.NewLine))
                    {
                        comment = comment.Replace(Environment.NewLine, " ");
                    }

                    int elapsedTimeInSeconds = 0;
                    if (elapsedTime != "")
                    {
                        elapsedTimeInSeconds = StringManipulation.ConvertTimespanStringToSeconds(elapsedTime);
                    }

					string milestoneName = "";
                    Milestone currentMilestone = listOfMilestones.Find(m => m.MilestoneID == currentRun.MilestoneID);

					if (!string.IsNullOrWhiteSpace(currentRun.MilestoneID))
					{
						milestoneName = currentMilestone.MilestoneName;
					}
					else
					{
						string milestoneString = "";
						if (projectName.Contains("Unity "))
						{
							milestoneString = projectName.Remove(0, 6);
						}
						milestoneName = milestoneString;

					}

					int completedInt = 0;

					if (completedDate == "")
					{
						completedInt = 0;
					}
					else
					{
						completedInt = Int32.Parse(completedDate);
					}

                    Test currentTest;
					currentTest.Area = area;
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
					currentTest.EditorVersion = AccessTestRail.GetEditorVersion(client, projectID, editorVersion);
					currentTest.Estimate = estimateInSeconds;
                    currentTest.EstimateForecast = estimateForecastInSeconds;
					currentTest.elapsedTimeInSeconds = elapsedTimeInSeconds;
					currentTest.CompletedDate = completedInt;
					currentTest.MilestoneID = currentRun.MilestoneID;
					currentTest.MilestoneName = milestoneName;
                    currentTest.identifier = caseID + "_" + testID;
					//currentTest.isRetest = isRetest.ToString();

                    listOfTests.Add(currentTest);
				}

			}
            List<Test> sortedList = SortListOfTests(listOfTests);

			string csvOfTests = CreateCSVOfTests(sortedList, listOfCases);
            Console.WriteLine(csvOfTests);

            //GoogleSheets.OutputTestsToGoogleSheets(sortedList, previousResults, listOfCases);

            Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}

        public static List<Case> CreateListOfCases(APIClient client, JArray caseTypes, JArray casesArray, List<Case> listOfCases, string suiteID, string suiteName, string milestoneName)
        {
            for (int i = 0; i < casesArray.Count; i++)
            {
                JObject arrayObject = casesArray[i].ToObject<JObject>();

                allCaseIDs.Add(Int32.Parse(arrayObject.Property("id").Value.ToString()));

                string caseID = arrayObject.Property("id").Value.ToString();
                string caseName = arrayObject.Property("title").Value.ToString();
                if (caseName.Contains(","))
                {
                    caseName = caseName.Replace(',', ' ');
                }
                if (caseName.Contains('"'))
                {
                    caseName = caseName.Replace('"', ' ');
                }
                string caseType = arrayObject.Property("type_id").Value.ToString();
                string sectionID = arrayObject.Property("section_id").Value.ToString();

                string createdOn = arrayObject.Property("created_on").Value.ToString();
                string updatedOn = arrayObject.Property("updated_on").Value.ToString();
				string estimate = "";
				string estimateForecast = "";

                string area = "";
                if (arrayObject.Property("area") != null && arrayObject.Property("area").Value != null && !string.IsNullOrWhiteSpace(arrayObject.Property("area").Value.ToString()))
                {
                    area = arrayObject.Property("area").Value.ToString();
                }
                //string area = arrayObject.Property("area").Value.ToString();

                JObject section = AccessTestRail.GetSection(client, sectionID);
                string sectionName = AccessTestRail.GetSectionName(section);

                if (sectionName.Contains(","))
                {
                    StringBuilder sb = new StringBuilder(sectionName);
                    sb[sectionName.IndexOf(',')] = ' ';
                    sectionName = sb.ToString();
                }

				int estimateInSeconds = 0;
				if (arrayObject.Property("estimate").Value != null && !string.IsNullOrWhiteSpace(arrayObject.Property("estimate").Value.ToString()))
                {
					estimate = arrayObject.Property("estimate").Value.ToString();
                    if (estimate != "")
                    {
                        estimateInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimate);
                    }
                }

                int estimateForecastInSeconds = 0;
				if (arrayObject.Property("estimate_forecast").Value != null && !string.IsNullOrWhiteSpace(arrayObject.Property("estimate_forecast").Value.ToString()))
                {
					estimateForecast = arrayObject.Property("estimate_forecast").Value.ToString();
                    if (estimateForecast != "")
                    {
                        estimateForecastInSeconds = StringManipulation.ConvertTimespanStringToSeconds(estimateForecast);
                    }
                }

				string caseIdentifier = caseName.Replace(" ", "").Replace(",", "") + "_" + sectionName.Replace(" ", "").Replace(",", "");

				//bool isRetest = listOfCases.Any(c => c.CaseID == Int32.Parse(caseID));

                

                Case newCase;
				newCase.Area = area;
                newCase.SuiteID = suiteID;
                newCase.SuiteName = suiteName;
                newCase.CreatedOn = createdOn;
                newCase.UpdatedOn = updatedOn;
                newCase.Section = sectionName;
                newCase.CaseID = Int32.Parse(caseID);
                newCase.CaseName = caseName;
                newCase.Status = StringManipulation.IsInvalid(arrayObject);
                newCase.Type = StringManipulation.GetCaseType(caseTypes, caseType);
                newCase.TemplateStatus = "";
				newCase.Estimate = estimateInSeconds;
				newCase.EstimateForecast = estimateForecastInSeconds;
				newCase.MilestoneName = milestoneName;
				newCase.UniqueCaseIdentifier = caseIdentifier;

                listOfCases.Add(newCase);
            }
            return listOfCases;
        }

        public static string CreateCSVOfTests(List<Test> sortedList, List<Case> listOfCases)
		{
            StringBuilder csv = new StringBuilder();
			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}", "Unique Case", "Milestone ID", "Milestone Name", "Area", "Suite Name", "Case ID", "Run Name", "Run ID", "Complete", "Identifier", "Section", "Title", "Created On", "Updated On", "Config", "Case Type", "Editor Version", "Completed Date", "Last Defects", "Last Comment", "Last Run Result", "Elapsed Time", "Estimate", "Forecast", "is Retest", "\n");
			csv.Append(header);
            List<int> passValues = new List<int>();
            for (int i = 0; i < sortedList.Count; i++)
            {
                Test testObject = sortedList[i];
                Case caseObject = listOfCases.Find(x => x.CaseID == testObject.CaseID); //finding the case that matches the test

                bool isRetest;

                Test findCase = sortedList.Find(o => o.CaseID == testObject.CaseID);
                if (findCase.Title != null)
                {
                    if (sortedList.IndexOf(findCase) < i)
                    {
                        isRetest = true;
                    }
                    else
                    {
                        isRetest = false;
                    }
                }
                else
                {
                    isRetest = false;
                }

                if (testObject.CaseID != 0)
                {
                    if (i != 0)
                    {
                        csv.Append("\n"); //removes the blank row between the headings and the first result
                    }
					string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},", caseObject.UniqueCaseIdentifier, testObject.MilestoneID, testObject.MilestoneName, testObject.Area, "\"" + testObject.SuiteName + "\"", testObject.CaseID, "\"" + testObject.RunName + "\"", testObject.RunID, testObject.isRunCompleted, testObject.identifier, "\"" + caseObject.Section + "\"", "\"" + testObject.Title + "\"", caseObject.CreatedOn, caseObject.UpdatedOn, testObject.Config, caseObject.Type, testObject.EditorVersion, testObject.CompletedDate, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", testObject.Status, testObject.elapsedTimeInSeconds.ToString(), testObject.Estimate, testObject.EstimateForecast, isRetest);

                    csv.Append(line);
                }

			}
            csv.Append("\n");
            csv.Append("\n");

            for (int k = 0; k < allCaseIDs.Count; k++)
            {
                if (!caseIDsInMilestone.Contains(allCaseIDs[k]))
                {
                    //List<Case> sortedListOfCases = SortListOfCases(listOfCases);

					Case caseNotRun = listOfCases.Find(x => x.CaseID == allCaseIDs[k]);

					string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}", caseNotRun.UniqueCaseIdentifier, "", caseNotRun.MilestoneName, caseNotRun.Area, "\"" + caseNotRun.SuiteName + "\"", caseNotRun.CaseID, "Not included in test run", "", "false", caseNotRun.CaseID + "_00", "\"" +  caseNotRun.Section + "\"", "\"" + caseNotRun.CaseName + "\"", caseNotRun.CreatedOn, caseNotRun.UpdatedOn, "", caseNotRun.Type, "Never Tested", "", "", "", "Untested", "0", caseNotRun.Estimate, caseNotRun.EstimateForecast, "FALSE", "\n");
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
			List<Test> sortedList = listOfTests.OrderBy(o => o.CompletedDate).ToList();
            return sortedList;
        }

        //private static List<Case> SortListOfCases(List<Case> listOfCases)
        //{
        //    List<Case> sortedList = listOfCases.OrderByDescending(o => o.CaseID).ToList();
        //    return sortedList;
        //}
	}
}
