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
        public static List<Case> listOfCases = new List<Case>();

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
        private static void EvaluateChoice(APIClient client)
        {
            Console.WriteLine("To create a CSV of all cases, press 1,");
            Console.WriteLine("To create a CSV of all tests in one long list (quicker but not organised), press 2,");
            Console.WriteLine("To create a CSV of all tests with orgnanised rows, press 3");

            string selection = Console.ReadLine();

            if (selection == "1")
            {
                GetAllCases(client, true);
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
                    GetAllCases(client, false);
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



        /// <summary>
        /// Retrieves TestRail cases using the API and puts them into a JArray so a CSV can be made
        /// </summary>
		private static void GetAllCases(APIClient client, bool fast)
		{
			JArray suitesArray = AccessTestRail.GetSuitesInProject(client, "2");

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

                //allCaseIDs.Add(id);

                JArray casesArray = AccessTestRail.GetCasesInSuite(client, "2", id);

                if (fast == true)
                {
                    string casesCSV = CreateCsvOfCases(client, casesArray, true);
                    Console.WriteLine(casesCSV);
                }
                else
                {
                    string casesCSV = CreateCsvOfCases(client, casesArray, false);
                }
			}

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
		}


        /// <summary>
        /// Retrieves TestRail tests (both in and out of plans)
        /// </summary>
        /// <param name="previousResults">Number of previous results to include.</param>
		private static void GetAllTests(APIClient client, int previousResults, bool simple)
		{
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();

            JArray c = AccessTestRail.GetRunsForMilestone(client, milestoneID);
            JArray planArray = AccessTestRail.GetPlansForMilestone(client, milestoneID);
			//The response includes an array of test plans. Each test plan in this list follows the same format as get_plan, except for the entries field which is not included in the response.


            AccessTestRail.GetSuitesAndRuns(c, suiteIDs, runIDs);

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
                ostrm = new FileStream("Tests"+ DateTime.UtcNow.ToFileTimeUtc().ToString() +".csv", FileMode.OpenOrCreate, FileAccess.Write);
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

                    if (!caseIDsInMilestone.Contains(caseID.ToString()))
                    {
                        caseIDsInMilestone.Add(caseID.ToString());
                    }

                    title = testObject.Property("title").Value.ToString();
                    status = StringManipulation.GetStatus(testObject.Property("status_id").Value.ToString());

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

            List<string> runInPlanIds = AccessTestRail.GetRunsInPlan(planArray, client, suiteInPlanIDs);

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
                    caseIDsInMilestone.Add(caseID.ToString());

					title = testObject.Property("title").Value.ToString();
                    status = StringManipulation.GetStatus(testObject.Property("status_id").Value.ToString());

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
                string csvOfTests = CreateCSVOfTestsComplete(sortedList, previousResults);
				Console.WriteLine(csvOfTests);
            }
            else
            {
                string csvOfTests = CreateCSVOfTestsOld(sortedList, previousResults);
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

            JArray c = AccessTestRail.GetRunsForMilestone(client, milestoneID);
            JArray planArray = AccessTestRail.GetPlansForMilestone(client, milestoneID);

            AccessTestRail.GetSuitesAndRuns(c, suiteIDs, runIDs);

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
                JArray testsArray = AccessTestRail.GetTestsInRun(client, runIDs[i]);

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

            List<string> runInPlanIds = AccessTestRail.GetRunsInPlan(planArray, client, suiteInPlanIDs);

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
				JArray testsArray = AccessTestRail.GetTestsInRun(client, runInPlanIds[i]);

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

        private static string CreateCsvOfCases(APIClient client, JArray casesArray, bool fast)
		{
			StringBuilder csv = new StringBuilder();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Case ID", "Suite ID", "Title", "References", "Case Status", "Steps", "Steps_Separated", "\n");
			csv.Append(header);

            for (int i = 0; i < casesArray.Count; i++)
            {
                JObject arrayObject = casesArray[i].ToObject<JObject>();



                allCaseIDs.Add(arrayObject.Property("id").Value.ToString());

                JObject suite = (JObject)client.SendGet($"get_suite/" + arrayObject.Property("suite_id").Value.ToString());
                string suiteName = suite.Property("name").Value.ToString();

                if (fast == false)
                {
                    string caseID = arrayObject.Property("id").Value.ToString();
                    string suiteID = arrayObject.Property("suite_id").Value.ToString();
                    string caseName = arrayObject.Property("title").Value.ToString();

                    Case newCase = new Case(suiteID, suiteName, caseID, caseName);
                    listOfCases.Add(newCase);
                }

                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", arrayObject.Property("id").Value.ToString(), arrayObject.Property("suite_id").Value.ToString(), "\"" + arrayObject.Property("title").Value.ToString() + "\"", "\"" + arrayObject.Property("refs").Value + "\"", StringManipulation.IsInvalid(arrayObject), StringManipulation.HasSteps(arrayObject), StringManipulation.HasStepsSeparated(arrayObject), "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

        public static string CreateCSVOfTestsComplete(List<Test> sortedList, int previousResults)
		{
			StringBuilder csv = new StringBuilder();
            string header = string.Format("{0},{1},{2},{3},{4},{5},", "Suite Name", "Title", "Last Run Result", "Previous Result", "Pass Rate", "\n");
			csv.Append(header);
			int count = 0;
            List<int> passValues = new List<int>();
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
                                string passRate = "";
								string line = string.Format("{0},", arrayObject.Status);
                                // 2) add the status to the same list
                                if (arrayObject.Status == "Passed")
                                {
                                    passValues.Add(100);
                                }
                                else
                                {
                                    passValues.Add(0);
                                }

                                csv.Append(line);
                                // if (count-1)=previousResults, calculate pass rate using the small list of pass values
                                if (count == (previousResults - 1))
                                {
                                    // eg sum(passvalues) / previousResults
                                    int sumOfValues = passValues.Sum();
                                    passRate = (sumOfValues / previousResults).ToString();
                                    csv.Append(string.Format("{0},", passRate + "%"));
                                }
								

							}
							else
							{
							}
						}
						else
						{
                            passValues.Clear();
							count = 0;
							csv.Append("\n");
							string line = string.Format("{0},{1},{2},", arrayObject.SuiteName, "\"" + arrayObject.Title + "\"", arrayObject.Status);
                            // 1) add the status to a list?
                            // if its a pass, value is 100
                            if (arrayObject.Status == "Passed")
                            {
                                passValues.Add(100);
                            }
                            else
                            {
                                passValues.Add(0);
                            }
							csv.Append(line);

						}
					}
				}
			}
            csv.Append("\n");

            for (int k = 0; k < allCaseIDs.Count; k++)
            {
                if (!caseIDsInMilestone.Contains(allCaseIDs[k]))
                {
                    List<Case> sortedListOfCases = SortListOfCases();

                    Case caseNotRun = sortedListOfCases.Find(x => x.CaseID == allCaseIDs[k]);

                    string line = string.Format("{0},{1},{2},{3},", caseNotRun.SuiteName, "\"" + caseNotRun.CaseName + "\"", "Has not been run", "\n");
                    csv.Append(line);
                }
            }

			return csv.ToString();
		}

        public static string CreateCSVOfTestsOld(List<Test> sortedList, int previousResults)
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

		public static string CreateCSVOfTestsFast(JArray arrayOfTests, string suiteId, string suiteName)
		{
			StringBuilder csv = new StringBuilder();

			string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Status", "Estimate", "\n");
			csv.Append(header);

			for (int i = 0; i < arrayOfTests.Count; i++)
			{
				JObject arrayObject = arrayOfTests[i].ToObject<JObject>();
                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", StringManipulation.GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

        //public static string CheckIfCaseHasBeenRunInMilestone(string currentCaseID)
        //{
        //    if (caseIDsInMilestone.Contains(currentCaseID))
        //    {
        //        return "Run";
        //    }
        //    else
        //    {
        //        return "Not run";
        //    }
        //}

        /// <summary>
        /// Sorts the list of tests by case_id and then run_id
        /// </summary>
        /// <returns>The list of tests.</returns>
        private static List<Test> SortListOfTests()
        {
            List<Test> sortedList = listOfTests.OrderBy(o => o.CaseID).ThenByDescending(o=>o.RunID).ToList();
            return sortedList;
        }

        private static List<Case> SortListOfCases()
        {
            List<Case> sortedList = listOfCases.OrderBy(o => o.CaseID).ToList();
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

public class Case
{
    public string SuiteID;
    public string SuiteName;
    public string CaseID;
    public string CaseName;

    public Case(string suiteID, string suiteName, string caseID, string caseName)
    {
        SuiteID = suiteID;
        SuiteName = suiteName;
        CaseID = caseID;
        CaseName = caseName;
    }
}
