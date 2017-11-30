using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Data;

namespace TestRailResultExport
{
	class MainClass
	{
		public static string milestoneID = "";
		public static List<string> suiteIDs = new List<string>();
		public static List<string> runIDs = new List<string>();
        public static List<string> allCaseIDs = new List<string>();

		public static List<string> suiteInPlanIDs = new List<string>();

        public static List<Test> listOfTests = new List<Test>();

		private static readonly IConfigReader _configReader = new ConfigReader();

		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			APIClient client = ConnectToTestrail();

			GetAllTests(client);


            //ConvertDataTableToCsv(sortedTable);
		}

		private static APIClient ConnectToTestrail()
		{
			APIClient client = new APIClient("http://qatestrail.hq.unity3d.com");
			client.User = _configReader.TestRailUser;
			client.Password = _configReader.TestRailPass;
			return client;
		}

		private static JArray GetRunsForMilestone(APIClient client, string milestoneID)
		{
			return (JArray)client.SendGet("get_runs/2&milestone_id=" + milestoneID);
		}

		private static JArray GetPlansForMilestone(APIClient client, string milestoneID)
		{
			return (JArray)client.SendGet("get_plans/2&milestone_id=" + milestoneID);
		}

		private static JArray GetTestsInRun(APIClient client, string runID)
		{
			return (JArray)client.SendGet("get_tests/" + runID);
		}

		private static JArray GetCasesInSuite(APIClient client, string projectID, string suiteID)
		{
			return (JArray)client.SendGet("get_cases/" + projectID + "&suite_id=" + suiteID);
		}

		private static JArray GetSuitesInProject(APIClient client, string projectID)
		{
			return (JArray)client.SendGet("get_suites/" + projectID);
		}

		private static JArray GetLatestResultsOfTest(APIClient client, string testID, string amountOfResultsToShow)
		{
			return (JArray)client.SendGet("get_results/" + testID + "&limit=" + amountOfResultsToShow);
		}

        private static JArray GetLatestResultsForCase(APIClient client, string runID, string caseID, string amountOfResultsToShow)
        {
          return (JArray)client.SendGet("get_results_for_case/" + runID + "/"+ caseID + "&limit=" + amountOfResultsToShow);
        }

		private static void GetAllCases(APIClient client)
		{
			JArray suitesArray = GetSuitesInProject(client, "2");

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
				ostrm = new FileStream("Cases.csv", FileMode.OpenOrCreate, FileAccess.Write);
				writer = new StreamWriter(ostrm);
			}
			catch (Exception e)
			{
				Console.WriteLine("Cannot open Cases.csv for writing");
				Console.WriteLine(e.Message);
				return;
			}
			Console.SetOut(writer);

			for (int i = 0; i < suitesArray.Count; i++)
			{
				JObject arrayObject = suitesArray[i].ToObject<JObject>();
				string id = arrayObject.Property("id").Value.ToString();

				JArray casesArray = GetCasesInSuite(client, "2", id);

				string casesCSV = CreateCsvOfCases(casesArray);
				Console.WriteLine(casesCSV);
			}

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}

		private static void GetAllTests(APIClient client)
		{
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();

			JArray c = GetRunsForMilestone(client, milestoneID);
			JArray planArray = GetPlansForMilestone(client, milestoneID);
			//The response includes an array of test plans. Each test plan in this list follows the same format as get_plan, except for the entries field which is not included in the response.


			GetSuitesAndRuns(c);

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
				ostrm = new FileStream("Tests.csv", FileMode.OpenOrCreate, FileAccess.Write);
				writer = new StreamWriter(ostrm);
			}
			catch (Exception e)
			{
				Console.WriteLine("Cannot open Tests.csv for writing");
				Console.WriteLine(e.Message);
				return;
			}
			Console.SetOut(writer);

			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Suite ID", "Suite Name", "Run ID", "Test ID", "CaseID", "Title", "Status", "\n");
			Console.WriteLine(header);

			for (int i = 0; i < runIDs.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runIDs[i].ToString());

                string testID = "";
                string caseID = "";
                string title = "";
                string status = "";

                for (int j = 0; j < testsArray.Count; j++)
                {
                    JObject testObject = testsArray[j].ToObject<JObject>();
                    testID = testObject.Property("id").Value.ToString();
                    caseID = testObject.Property("case_id").Value.ToString();
                    title = testObject.Property("title").Value.ToString();
                    status = GetStatus(testObject.Property("status_id").Value.ToString());

                    allCaseIDs.Add(caseID);

					string suiteName = "";

					if (suiteIDs[i] != "0")
					{
						JObject suite = (JObject)client.SendGet($"get_suite/{suiteIDs[i]}");
						suiteName = suite.Property("name").Value.ToString();
					}
					else
					{
						suiteName = "deleted";
					}

					Test currentTest = new Test(suiteIDs[i], suiteName, runIDs[i], testID, caseID, title, status);
					listOfTests.Add(currentTest);
                }

				

				//string csvOfTests = CreateCSVOfTests();
				//Console.WriteLine(csvOfTests);

			}

			List<string> runInPlanIds = GetRunsInPlan(planArray, client);

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runInPlanIds[i].ToString());

				string testID = "";
				string caseID = "";
				string title = "";
				string status = "";

				for (int j = 0; j < testsArray.Count; j++)
				{
					JObject testObject = testsArray[j].ToObject<JObject>();

                    testID = testObject.Property("id").Value.ToString();
					caseID = testObject.Property("case_id").Value.ToString();
					title = testObject.Property("title").Value.ToString();
					status = GetStatus(testObject.Property("status_id").Value.ToString());

					allCaseIDs.Add(caseID);

					string suiteName = "";

					if (suiteInPlanIDs[i] != "0")
					{
						JObject suite = (JObject)client.SendGet($"get_suite/{suiteInPlanIDs[i]}");
						suiteName = suite.Property("name").Value.ToString();
					}
					else
					{
						suiteName = "deleted";
					}

					Test currentTest = new Test(suiteInPlanIDs[i], suiteName, runInPlanIds[i], testID, caseID, title, status);
					listOfTests.Add(currentTest);
				}

			}
            List<Test> sortedList = SortListOfTests();
            string csvOfTests = CreateCSVOfTests(sortedList);
			Console.WriteLine(csvOfTests);

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}


		private static string GetStatus(string rawValue)
		{
			if (rawValue == "1")
			{
				return "Passed";
			}
			else if (rawValue == "2")
			{
				return "Blocked";
			}
			else if (rawValue == "3")
			{
				return "Untested";
			}
			else if (rawValue == "4")
			{
				return "Retest";
			}
			else if (rawValue == "5")
			{
				return "Failed";
			}
			else
			{
				return "Other";
			}
		}

		private static List<int> GetAllSuites(JArray arrayOfSuites)
		{
			List<int> listOfSuiteIds = new List<int>();
			for (int i = 0; i < arrayOfSuites.Count; i++)
			{
				JObject arrayObject = arrayOfSuites[i].ToObject<JObject>();
				int id = Int32.Parse(arrayObject.Property("id").Value.ToString());
				listOfSuiteIds.Add(id);
			}
			return listOfSuiteIds;
		}

		private static List<string> GetRunsInPlan(JArray planArray, APIClient client)
		{
			List<JArray> ListOfRunsInPlan = new List<JArray>();
			List<string> planIds = new List<string>();
			List<string> runInPlanIds = new List<string>();

			for (int i = 0; i < planArray.Count; i++)
			{
				JObject arrayObject = planArray[i].ToObject<JObject>();

				string planID = arrayObject.Property("id").Value.ToString();
				planIds.Add(planID);

				foreach (string id in planIds)
				{
					JObject singularPlanObject = (JObject)client.SendGet("get_plan/" + id);

					JProperty prop = singularPlanObject.Property("entries");
					if (prop != null && prop.Value != null)
					{
						JArray entries = (JArray)singularPlanObject.Property("entries").First;

						for (int k = 0; k < entries.Count; k++)
						{
							JObject entriesObject = entries[k].ToObject<JObject>();


							JArray runsArray = (JArray)entriesObject.Property("runs").First;

							for (int j = 0; j < runsArray.Count; j++)
							{
								JObject runObject = runsArray[j].ToObject<JObject>();


								string runInPlanId = runObject.Property("id").Value.ToString();

								if (!runInPlanIds.Contains(runInPlanId))
								{
									runInPlanIds.Add(runInPlanId);
									string suiteInPlanId = runObject.Property("suite_id").Value.ToString();
									suiteInPlanIDs.Add(suiteInPlanId);
								}
							}
						}
					}
				}
			}
			return runInPlanIds;
		}

		public static void GetSuitesAndRuns(JArray runsArr)
		{
			for (int i = 0; i < runsArr.Count; i++)
			{
				JObject arrayObject = runsArr[i].ToObject<JObject>();

				JProperty prop = arrayObject.Property("suite_id");
				if (prop != null && prop.Value != null && !string.IsNullOrEmpty(prop.Value.ToString()))
				{
					string suite_id = arrayObject.Property("suite_id").Value.ToString();

					string run_id = arrayObject.Property("id").Value.ToString();

					suiteIDs.Add(suite_id);
					runIDs.Add(run_id);
				}
				else
				{
					string suite_id = "0";

					string run_id = arrayObject.Property("id").Value.ToString();

					suiteIDs.Add(suite_id);
					runIDs.Add(run_id);
				}
			}
		}

		private static string CreateCsvOfCases(JArray casesArray)
		{
			StringBuilder csv = new StringBuilder();

			string header = string.Format("{0},{1},{2},{3},{4}", "Case ID", "Suite ID", "Title", "Milestone ID", "\n");
			csv.Append(header);

			for (int i = 0; i < casesArray.Count; i++)
			{
				JObject arrayObject = casesArray[i].ToObject<JObject>();

				string newLine = string.Format("{0},{1},{2},{3},{4}", arrayObject.Property("id").Value, arrayObject.Property("suite_id").Value, arrayObject.Property("title").Value, arrayObject.Property("milestone_id").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

		public static string CreateCSVOfRuns(JArray rawData)
		{
			StringBuilder csv = new StringBuilder();

			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", "Run ID", "Suite ID", "Title", "Is_Completed", "Passed_count", "Failed_count", "Blocked_count", "pending_count", "Untested_count", "Milestone ID", "url", "\n");
			csv.Append(header);

			for (int i = 0; i < rawData.Count; i++)
			{
				JObject arrayObject = rawData[i].ToObject<JObject>();

				string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", arrayObject.Property("id").Value, arrayObject.Property("suite_id").Value, arrayObject.Property("name").Value, arrayObject.Property("is_completed").Value, arrayObject.Property("passed_count").Value, arrayObject.Property("failed_count").Value, arrayObject.Property("blocked_count").Value, arrayObject.Property("custom_status1_count").Value, arrayObject.Property("untested_count").Value, arrayObject.Property("milestone_id").Value, arrayObject.Property("url").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

		public static string CreateCSVOfPlanRuns(JArray rawData)
		{
			StringBuilder csv = new StringBuilder();

			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", "Plan ID", "Suite ID", "Title", "Is_Completed", "Passed_count", "Failed_count", "Blocked_count", "pending_count", "Untested_count", "Milestone ID", "url", "\n");

			for (int i = 0; i < rawData.Count; i++)
			{
				JObject arrayObject = rawData[i].ToObject<JObject>();

				string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", arrayObject.Property("plan_id").Value, arrayObject.Property("suite_id").Value, arrayObject.Property("name").Value, arrayObject.Property("is_completed").Value, arrayObject.Property("passed_count").Value, arrayObject.Property("failed_count").Value, arrayObject.Property("blocked_count").Value, arrayObject.Property("custom_status1_count").Value, arrayObject.Property("untested_count").Value, arrayObject.Property("milestone_id").Value, arrayObject.Property("url").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

        public static string CreateCSVOfTests(List<Test> sortedList)
		{
			StringBuilder csv = new StringBuilder();

            //string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Suite ID", "Suite Name", "Run ID", "Test ID", "CaseID", "Title", "Status", "Estimate", "\n");
            //csv.Append(header);

            //List<Test> sortedList = SortListOfTests();

			for (int i = 0; i < sortedList.Count; i++)
			{
				Test arrayObject = sortedList[i];
                if (i != 0)
                {
                    if (arrayObject.CaseID != "")
                    {
                        if (arrayObject.CaseID == sortedList[i - 1].CaseID)
                        {
							string line = string.Format("{0},{1},{2},", arrayObject.RunID, arrayObject.TestID, arrayObject.Status);
							csv.Append(line);
                        }
                        else
                        {
                            csv.Append("\n");
                            string line = string.Format("{0},{1},{2},{3},{4},{5},{6}", arrayObject.SuiteID, arrayObject.SuiteName, arrayObject.RunID, arrayObject.TestID, arrayObject.CaseID, "\"" + arrayObject.Title + "\"", arrayObject.Status);
                            csv.Append(line);
                        }
                    }
                }
			}

			return csv.ToString();
		}


		//public static string CreateCSVOfTestsWithResults(JArray arrayOfTests, int suiteId, string suiteName)
		//{
		//	StringBuilder csv = new StringBuilder();

		//	string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Status", "\n");
		//	csv.Append(header);

		//	for (int j = 0; j < arrayOfTests.Count; j++)
		//	{
		//		JObject testObject = arrayOfTests[j].ToObject<JObject>();
		//		string caseID = testObject.Property("case_id").Value.ToString();

		//		allCaseIDs.Add(caseID);
		//	}

  //          List<string> duplicates = FindDupesInCaseIDs();
  //          //if the case id is in the duplicates list, add all reults to same line

		//	for (int i = 0; i < arrayOfTests.Count; i++)
		//	{
		//		JObject arrayObject = arrayOfTests[i].ToObject<JObject>();

		//		string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value);
  //              csv.Append(newLine);

		//		string currentCaseID = arrayObject.Property("case_id").Value.ToString();

  //              if (duplicates.Contains(currentCaseID))
  //              {
  //                  foreach (string caseDupe in duplicates)
  //                  {
  //                      if (caseDupe == currentCaseID)
  //                      {
  //                          //get results for test id
  //                          //JArray resultsOfTest = 
  //                      }
  //                  }
  //              }
  //              else
  //              {
  //                  csv.Append("\n");
  //              }

		//		//string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value, "\n");
				
		//	}

		//	return csv.ToString();
		//}



		//public static string CreateCSVOfTestsWithMostRecentResults(APIClient client, JArray arrayOfTests, int suiteId, string suiteName, string numberOfResultsToGet)
		//{
		//	StringBuilder csv = new StringBuilder();

		//	string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Estimate", "Created on", "defects", "elapsed", "status", "comment", "\n");
		//	csv.Append(header);

		//	for (int i = 0; i < arrayOfTests.Count; i++)
		//	{
		//		JObject arrayObject = arrayOfTests[i].ToObject<JObject>();

		//		string testID = arrayObject.Property("id").Value.ToString();

		//		JArray resultsArray = GetLatestResultsOfTest(client, testID, numberOfResultsToGet);

		//		for (int j = 0; j < resultsArray.Count; j++)
		//		{
		//			JObject resultsObject = resultsArray[j].ToObject<JObject>();
		//			DateTime createdOnDate = Convert.ToDateTime(resultsObject.Property("created_on").Value.ToString());


		//			string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", arrayObject.Property("estimate").Value, createdOnDate, "\"" + resultsObject.Property("defects") + "\"", resultsObject.Property("elapsed"), GetStatus(resultsObject.Property("status_id").Value.ToString()), resultsObject.Property("comment"), "\n");
		//			csv.Append(newLine);

		//		}
		//	}

		//	return csv.ToString();
		//}

        private static List<Test> SortListOfTests()
        {
            List<Test> sortedList = listOfTests.OrderBy(o => o.CaseID).ToList();
            return sortedList;
        }
	}
}

public class Test
{
    public string SuiteID;
    public string SuiteName;
    public string RunID;
    public string TestID;
    public string CaseID;
    public string Title;
    public string Status;

    public Test(string suiteID, string suiteName, string runID, string testID, string caseID, string title, string status)
    {
        SuiteID = suiteID;
        SuiteName = suiteName;
        RunID = runID;
        TestID = testID;
        CaseID = caseID;
        Title = title;
        Status = status;
    }
}