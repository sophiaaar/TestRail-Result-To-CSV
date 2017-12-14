﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Data;
using System.Threading;

namespace TestRailResultExport
{
	class MainClass
	{
		public static string milestoneID = "";
		public static List<string> suiteIDs = new List<string>();
		public static List<string> runIDs = new List<string>();
        public static List<string> allCaseIDs = new List<string>();
        public static List<string> caseIDsInMilestone = new List<string>(); //case IDs that have been run

        public static int numberPassed;
        public static int numberFailed;
        public static int numberBlocked;

		public static List<string> suiteInPlanIDs = new List<string>();

        public static List<Test> listOfTests = new List<Test>();

		private static readonly IConfigReader _configReader = new ConfigReader();

		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			APIClient client = ConnectToTestrail();
            //GoogleSheets.ConnectToGoogleSheets();

            EvaluateChoice(client);
		}

		private static APIClient ConnectToTestrail()
		{
			APIClient client = new APIClient("http://qatestrail.hq.unity3d.com");
			client.User = _configReader.TestRailUser;
			client.Password = _configReader.TestRailPass;
			return client;
		}

        /// <summary>
        /// Asks the user which kind of action they'd like to perform, then calls the appropriate method
        /// </summary>
        /// <param name="client">Client.</param>
        private static void EvaluateChoice(APIClient client)
        {
            Console.WriteLine("To create a CSV of all cases, press 1,");
            Console.WriteLine("To create a CSV of all tests in one long list (quicker but not organised), press 2,");
            Console.WriteLine("To create a CSV of all tests with orgnanised rows, press 3");

            string selection = Console.ReadLine();

            if (selection == "1")
            {
                GetAllCases(client);
            }
            else if (selection == "2")
            {
                GetAllTestsFast(client);
            }
            else if (selection == "3")
            {
                Console.WriteLine("How many previous results do you want to see?");
                int previousResults = Int32.Parse(Console.ReadLine());

                Console.WriteLine("Simple view? Y or N");
                string simple = Console.ReadLine();

                if (simple == "Y" || simple == "y")
                {
                    GetAllTests(client, previousResults, true);
                }
                else
                {
                    GetAllTests(client, previousResults, false);
                }
            }
            else
            {
                Console.WriteLine("Please make a valid selection");
                EvaluateChoice(client);
            }
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

        /// <summary>
        /// Retrieves TestRail cases using the API and puts them into a JArray so a CSV can be made
        /// </summary>
        /// <param name="client">Client.</param>
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

                allCaseIDs.Add(id);

				JArray casesArray = GetCasesInSuite(client, "2", id);

				string casesCSV = CreateCsvOfCases(casesArray);
				Console.WriteLine(casesCSV);
			}

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}


        /// <summary>
        /// Retrieves TestRail tests (both in and out of plans)
        /// </summary>
        /// <param name="client">Client.</param>
        /// <param name="previousResults">Number of previous results to include.</param>
		private static void GetAllTests(APIClient client, int previousResults, bool simple)
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

			//string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Suite ID", "Suite Name", "Run ID", "Test ID", "CaseID", "Title", "Status", "\n");
			//Console.WriteLine(header);

			for (int i = 0; i < runIDs.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runIDs[i]);

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

                    caseIDsInMilestone.Add(caseID.ToString());

                    title = testObject.Property("title").Value.ToString();
                    status = GetStatus(testObject.Property("status_id").Value.ToString());

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

                    //allCaseIDs.Add(caseID);

					string suiteName = "";

                    // Some suites have been deleted, but the tests and runs remain
					if (suiteIDs[i] != "0")
					{
                        // Get the suite_id that corresponds to the run_id
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

			}

			List<string> runInPlanIds = GetRunsInPlan(planArray, client);

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runInPlanIds[i]);

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
                    caseIDsInMilestone.Add(caseID.ToString());

					title = testObject.Property("title").Value.ToString();
					status = GetStatus(testObject.Property("status_id").Value.ToString());

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

					//allCaseIDs.Add(caseID);

					string suiteName = "";

					// Some suites have been deleted, but the tests and runs remain
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

            Console.WriteLine("Number Passed, Number Failed, Number Blocked,\n");
            Console.WriteLine(string.Format("{0},{1},{2},{3},{4}", numberPassed, numberFailed, numberBlocked, "\n", "\n"));

            if (simple == true)
            {
                string csvOfTests = CreateCSVOfTestsSimpleView(sortedList, previousResults);
				Console.WriteLine(csvOfTests);
            }
            else
            {
                string csvOfTests = CreateCSVOfTests(sortedList, previousResults);
                Console.WriteLine(csvOfTests);
            }

            //GoogleSheets.OutputTestsToGoogleSheets(sortedList, previousResults);


			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}



        /// <summary>
        /// Retrieves all tests but doesn't order them
        /// </summary>
        /// <param name="client">APIClient.</param>
		private static void GetAllTestsFast(APIClient client)
		{
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();

			JArray c = GetRunsForMilestone(client, milestoneID);
			JArray planArray = GetPlansForMilestone(client, milestoneID);

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

			for (int i = 0; i < runIDs.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runIDs[i]);

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

				string csvOfTests = CreateCSVOfTestsFast(testsArray, suiteIDs[i], suiteName);
				Console.WriteLine(csvOfTests);

			}

			List<string> runInPlanIds = GetRunsInPlan(planArray, client);

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
				JArray testsArray = GetTestsInRun(client, runInPlanIds[i]);

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

				string csvOfTests = CreateCSVOfTestsFast(testsArray, suiteInPlanIDs[i], suiteName);
				Console.WriteLine(csvOfTests);
			}

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}

        /// <summary>
        /// Converts the status number to a string
        /// </summary>
        /// <returns>The status.</returns>
        /// <param name="rawValue">Raw value.</param>
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

        private static string HasSteps(JObject arrayObject)
        {
            if (arrayObject.Property("custom_steps") != null && !string.IsNullOrWhiteSpace(arrayObject.Property("custom_steps").Value.ToString()))
            {
                return "Yes";
            }
            else
            {
                return "No";
            }
        }

		private static string HasStepsSeparated(JObject arrayObject)
		{
            if (arrayObject.Property("custom_steps_separated") != null && !string.IsNullOrWhiteSpace(arrayObject.Property("custom_steps_separated").Value.ToString()))
			{
				return "Yes";
			}
			else
			{
				return "No";
			}
		}

        private static string IsInvalid(JObject arrayObject)
        {
            if (HasSteps(arrayObject) == "No" && HasStepsSeparated(arrayObject) == "No")
            {
                return "Invalid";
            }
            else
            {
                return "";
            }
        }

		private static string CreateCsvOfCases(JArray casesArray)
		{
			StringBuilder csv = new StringBuilder();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Case ID", "Suite ID", "Title", "References", "Case Status", "Steps", "Steps_Separated", "\n");
			csv.Append(header);

			for (int i = 0; i < casesArray.Count; i++)
			{
				JObject arrayObject = casesArray[i].ToObject<JObject>();

                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", arrayObject.Property("id").Value, arrayObject.Property("suite_id").Value, "\"" + arrayObject.Property("title").Value + "\"", "\"" + arrayObject.Property("refs").Value + "\"", IsInvalid(arrayObject), HasSteps(arrayObject), HasStepsSeparated(arrayObject), "\n");
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

        public static string CreateCSVOfTests(List<Test> sortedList, int previousResults)
		{
			StringBuilder csv = new StringBuilder();
			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Suite ID", "Suite Name", "Run ID", "Test ID", "CaseID", "Title", "Status", "\n");
            csv.Append(header);
            int count = 0;
			for (int i = 0; i < sortedList.Count; i++)
			{
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
                                string line = string.Format("{0},{1},{2},", arrayObject.RunID, arrayObject.TestID, arrayObject.Status);
                                csv.Append(line);
                            }
                            else
                            {
                            }
                        }
                        else
                        {
                            count = 0;
                            csv.Append("\n");
                            string line = string.Format("{0},{1},{2},{3},{4},{5},{6},", arrayObject.SuiteID, arrayObject.SuiteName, arrayObject.RunID, arrayObject.TestID, arrayObject.CaseID, "\"" + arrayObject.Title + "\"", arrayObject.Status);
                            csv.Append(line);

                        }
                    }
                }
			}

			return csv.ToString();
		}

		public static string CreateCSVOfTestsSimpleView(List<Test> sortedList, int previousResults)
		{
			StringBuilder csv = new StringBuilder();
			string header = string.Format("{0},{1},{2},{3},{4},", "Suite Name", "Title", "Last Run Result", "Previous Result", "\n");
			csv.Append(header);
			int count = 0;
			for (int i = 0; i < sortedList.Count; i++)
			{
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
								string line = string.Format("{0},", arrayObject.Status);
								csv.Append(line);
							}
							else
							{
							}
						}
						else
						{
							count = 0;
							csv.Append("\n");
							string line = string.Format("{0},{1},{2},", arrayObject.SuiteName, "\"" + arrayObject.Title + "\"", arrayObject.Status);
							csv.Append(line);

						}
					}
				}
			}

			return csv.ToString();
		}

		public static string CreateCSVOfTestsFast(JArray arrayOfTests, string suiteId, string suiteName)
		{
			StringBuilder csv = new StringBuilder();

			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Status", "Estimate", "\n");
			csv.Append(header);

			for (int i = 0; i < arrayOfTests.Count; i++)
			{
				JObject arrayObject = arrayOfTests[i].ToObject<JObject>();
				string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

        public static string CheckIfCaseHasBeenRunInMilestone(string currentCaseID)
        {
            if (caseIDsInMilestone.Contains(currentCaseID))
            {
                return "Run";
            }
            else
            {
                return "Not run";
            }
        }

        /// <summary>
        /// Sorts the list of tests by case_id and then run_id
        /// </summary>
        /// <returns>The list of tests.</returns>
        private static List<Test> SortListOfTests()
        {
            List<Test> sortedList = listOfTests.OrderBy(o => o.CaseID).ThenByDescending(o=>o.RunID).ToList();
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
    public int CaseID;
    public string Title;
    public string Status;

    public Test(string suiteID, string suiteName, string runID, string testID, int caseID, string title, string status)
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
