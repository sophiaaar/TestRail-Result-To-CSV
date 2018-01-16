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
	public class MainClass
	{
		public static string milestoneID = "";
		public static List<string> suiteIDs = new List<string>();
        public static List<string> suiteNames = new List<string>();
        public static List<string> suiteInPlanIDs = new List<string>();
		public static List<string> runIDs = new List<string>();
        public static List<Run> runs = new List<Run>();
        public static List<int> allCaseIDs = new List<int>();
        public static List<int> caseIDsInMilestone = new List<int>(); //case IDs that have been run

        public static int numberPassed;
        public static int numberFailed;
        public static int numberBlocked;

		private static readonly IConfigReader _configReader = new ConfigReader();

        public struct Test
        {
            public string SuiteID;
            public string SuiteName;
            public int RunID;
            public string TestID;
            public int CaseID;
            public string Title;
            public string Status;
            public string Defects;
            public string Comment;
            public string Config;
            public string EditorVersion;
        }

        public struct Case
        {
            public string SuiteID;
            public string SuiteName;
            public int CaseID;
            public string CaseName;
            public string Status;
            public string Type;
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
        }

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
            Console.WriteLine("To create a CSV of all tests, press 2");

            string selection = Console.ReadLine();

            if (selection == "1")
            {
                //JArray test = AccessTestRail.GetStatuses(client);
                GetAllCases(client, true);
            }
            else if (selection == "2")
            {
                Console.WriteLine("How many previous results do you want to see?");
                int previousResults = Int32.Parse(Console.ReadLine());

                GetAllTests(client, previousResults);
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
                string suiteName = arrayObject.Property("name").Value.ToString();


                JArray casesArray = AccessTestRail.GetCasesInSuite(client, "2", id);

                if (fast == true)
                {
                    string casesCSV = CreateCsvOfCases(casesArray, suiteName);
                    Console.WriteLine(casesCSV);
                }
                else
                {
                    string casesCSV = CreateCsvOfCases(casesArray, suiteName);
                    Console.WriteLine(casesCSV);
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
		private static void GetAllTests(APIClient client, int previousResults)
		{
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();

            JArray c = AccessTestRail.GetRunsForMilestone(client, milestoneID);
            JArray planArray = AccessTestRail.GetPlansForMilestone(client, milestoneID);
            //The response includes an array of test plans. Each test plan in this list follows the same format as get_plan, except for the entries field which is not included in the response.

            List<string> runInPlanIds = AccessTestRail.GetRunsInPlan(planArray, client, suiteInPlanIDs, runs);

            List<Case> listOfCases = new List<Case>();
            List<Test> listOfTests = new List<Test>();
            List<Suite> listOfSuites = new List<Suite>();

            JArray suitesArray = AccessTestRail.GetSuitesInProject(client, "2");

            for (int i = 0; i < suitesArray.Count; i++)
            {
                JObject arrayObject = suitesArray[i].ToObject<JObject>();
                string id = arrayObject.Property("id").Value.ToString();
                string suiteName = arrayObject.Property("name").Value.ToString(); //create list of suiteNames to use later

                Suite newSuite;
                newSuite.SuiteID = id;
                newSuite.SuiteName = suiteName;
                listOfSuites.Add(newSuite);


                JArray casesArray = AccessTestRail.GetCasesInSuite(client, "2", id);
                listOfCases = CreateListOfCases(casesArray, listOfCases, id, suiteName);
            }


            AccessTestRail.GetSuitesAndRuns(c, suiteIDs, runIDs, runs);

			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;

			try
			{
                ostrm = new FileStream("Tests"+ DateTime.UtcNow.ToLongDateString() +".csv", FileMode.OpenOrCreate, FileAccess.Write);
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

                //Test[] arrayOfTests = new Test[testsArray.Count];

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

                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
                    }

                    // Find config for runID
                    string runID = runs[i].RunID;
                    string config = runs[i].Config;

                    //append to csv at this point?

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

                    Test currentTest;
                    currentTest.SuiteID = suiteIDs[i];
                    currentTest.SuiteName = suiteName;
                    currentTest.RunID = Int32.Parse(runIDs[i]);
                    currentTest.TestID = testID;
                    currentTest.CaseID = caseID;
                    currentTest.Title = title;
                    currentTest.Status = status;
                    currentTest.Defects = defects;
                    currentTest.Comment = comment;
                    currentTest.Config = config;
                    currentTest.EditorVersion = StringManipulation.GetEditorVersion(editorVersion);

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

                    JArray resultsOfLatestTest = AccessTestRail.GetLatestResultsOfTest(client, testID, "1");

                    for (int k = 0; k < resultsOfLatestTest.Count; k++)
                    {
                        JObject resultObject = resultsOfLatestTest[k].ToObject<JObject>();

                        defects = resultObject.Property("defects").Value.ToString();
                        comment = resultObject.Property("comment").Value.ToString();
                        editorVersion = resultObject.Property("custom_editorversion").Value.ToString();
                    }

                    // Find config for runID
                    Run currentRun = runs.Find(o => o.RunID == runInPlanIds[i]);
                    string config = currentRun.Config;
                    string runID = currentRun.RunID;

                    if (config.Contains(','))
                    {
                        int index = config.IndexOf(',');
                        config = config.Substring(0, index);
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
                    else if (comment.Contains(Environment.NewLine))
                    {
                        comment = comment.Replace(Environment.NewLine, " ");
                    }

                    Test currentTest;
                    currentTest.SuiteID = suiteInPlanIDs[i];
                    currentTest.SuiteName = suiteName;
                    currentTest.RunID = Int32.Parse(runInPlanIds[i]);
                    currentTest.TestID = testID;
                    currentTest.CaseID = caseID;
                    currentTest.Title = title;
                    currentTest.Status = status;
                    currentTest.Defects = defects;
                    currentTest.Comment = comment;
                    currentTest.Config = config;
                    currentTest.EditorVersion = StringManipulation.GetEditorVersion(editorVersion);

					listOfTests.Add(currentTest);
				}

			}
            List<Test> sortedList = SortListOfTests(listOfTests);

            Console.WriteLine("Number Passed, Number Failed, Number Blocked,");
            Console.WriteLine(string.Format("{0},{1},{2},{3},{4}", numberPassed, numberFailed, numberBlocked, "\n", "\n"));

            string csvOfTests = CreateCSVOfTestsComplete(sortedList, previousResults, listOfCases);
            Console.WriteLine(csvOfTests);

            //GoogleSheets.OutputTestsToGoogleSheets(sortedList, previousResults);


			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();
			Console.WriteLine("Done");
		}

        private static string CreateCsvOfCases(JArray casesArray, string suiteName)
		{
			StringBuilder csv = new StringBuilder();

            List<Case> listOfCases = new List<Case>();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "Case ID", "Suite ID", "Title", "References", "Case Status", "Steps", "Steps_Separated", "\n");
			csv.Append(header);

            for (int i = 0; i < casesArray.Count; i++)
            {
                JObject arrayObject = casesArray[i].ToObject<JObject>();



                allCaseIDs.Add(Int32.Parse(arrayObject.Property("id").Value.ToString()));

                string caseID = arrayObject.Property("id").Value.ToString();
                string suiteID = arrayObject.Property("suite_id").Value.ToString();
                string caseName = arrayObject.Property("title").Value.ToString();
                string caseType = arrayObject.Property("type_id").Value.ToString();

                Case newCase;
                newCase.SuiteID = suiteID;
                newCase.SuiteName = suiteName;
                newCase.CaseID = Int32.Parse(caseID);
                newCase.CaseName = caseName;
                newCase.Status = StringManipulation.IsInvalid(arrayObject);
                newCase.Type = StringManipulation.GetCaseType(caseType);


                listOfCases.Add(newCase);

                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", arrayObject.Property("id").Value.ToString(), arrayObject.Property("suite_id").Value.ToString(), "\"" + arrayObject.Property("title").Value.ToString() + "\"", "\"" + arrayObject.Property("refs").Value + "\"", StringManipulation.IsInvalid(arrayObject), StringManipulation.HasSteps(arrayObject), StringManipulation.HasStepsSeparated(arrayObject), "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}

        public static List<Case> CreateListOfCases(JArray casesArray, List<Case> listOfCases, string suiteID, string suiteName)
        {
            for (int i = 0; i < casesArray.Count; i++)
            {
                JObject arrayObject = casesArray[i].ToObject<JObject>();

                allCaseIDs.Add(Int32.Parse(arrayObject.Property("id").Value.ToString()));

                string caseID = arrayObject.Property("id").Value.ToString();
                string caseName = arrayObject.Property("title").Value.ToString();
                string caseType = arrayObject.Property("type_id").Value.ToString();

                Case newCase;
                newCase.SuiteID = suiteID;
                newCase.SuiteName = suiteName;
                newCase.CaseID = Int32.Parse(caseID);
                newCase.CaseName = caseName;
                newCase.Status = StringManipulation.IsInvalid(arrayObject);
                newCase.Type = StringManipulation.GetCaseType(caseType);

                listOfCases.Add(newCase);
            }
            return listOfCases;
        }

        public static string CreateCSVOfTestsComplete(List<Test> sortedList, int previousResults, List<Case> listOfCases)
		{
			StringBuilder csv = new StringBuilder();
            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", "Suite Name", "Title", "Config", "Case Type", "Editor Version", "Last Defects", "Last Comment", "Last Run Result", "Previous Result", "Previous Result", "Pass Rate", "\n");
			csv.Append(header);
			int count = 0;
            List<int> passValues = new List<int>();
            for (int i = 0; i < sortedList.Count; i++)
            {
                Test testObject = sortedList[i];
                Case caseObject = listOfCases.Find(x => x.CaseID == testObject.CaseID); //finding the case that matches the test

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
                                string line = string.Format("{0},", testObject.Status);
                                // 2) add the status to the same list
                                if (testObject.Status == "Passed")
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
                        }
                        else
                        {
                            // Some values get reset here because this is a brand new line and a new case
                            passValues.Clear();
                            count = 0;
                            if (i != 0)
                            {
                                csv.Append("\n"); //removes the blank row between the headings and the first result
                            }
                            string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                            // 1) add the status to a list?
                            // if its a pass, value is 100
                            if (testObject.Status == "Passed")
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
                    else
                    {
                        // Some values get reset here because this is a brand new line and a new case
                        passValues.Clear();
                        count = 0;
                        if (i != 0)
                        {
                            csv.Append("\n"); //removes the blank row between the headings and the first result
                        }
                        string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                        // 1) add the status to a list?
                        // if its a pass, value is 100
                        if (testObject.Status == "Passed")
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

    //            if (i != 0)
    //            {
                    


				//}
                //else
                //{
                //    //// Some values get reset here because this is a brand new line and a new case
                //    //passValues.Clear();
                //    //count = 0;
                //    //if (i != 1)
                //    //{
                //    //    csv.Append("\n"); //removes the blank row between the headings and the first result
                //    //}
                //    //string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},", "\"" + testObject.SuiteName + "\"", "\"" + testObject.Title + "\"", "\"" + testObject.Config + "\"", "\"" + caseObject.Type + "\"", testObject.EditorVersion, "\"" + testObject.Defects + "\"", "\"" + testObject.Comment + "\"", "\"" + testObject.Status + "\"");
                //    //// 1) add the status to a list?
                //    //// if its a pass, value is 100
                //    //if (testObject.Status == "Passed")
                //    //{
                //    //    passValues.Add(100);
                //    //}
                //    //else
                //    //{
                //    //    passValues.Add(0);
                //    //}
                //    //csv.Append(line);
                //}
			}
            csv.Append("\n");
            csv.Append("\n");

            for (int k = 0; k < allCaseIDs.Count; k++)
            {
                if (!caseIDsInMilestone.Contains(allCaseIDs[k]))
                {
                    List<Case> sortedListOfCases = SortListOfCases(listOfCases);

                    Case caseNotRun = sortedListOfCases.Find(x => x.CaseID == allCaseIDs[k]);

                    string line = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", caseNotRun.SuiteName, "\"" + caseNotRun.CaseName + "\"", "", "\"" + caseNotRun.Type + "\"", "", "", "", "Untested", "\n");
                    csv.Append(line);
                }
                else
                {
                    //do this whole section first?
                }
            }

			return csv.ToString();
		}

        /// <summary>
        /// Sorts the list of tests by case_id and then run_id
        /// </summary>
        private static List<Test> SortListOfTests(List<Test> listOfTests)
        {
            //List<Test> sortedList = listOfTests.OrderByDescending(o => o.CaseID).ThenByDescending(o=>o.RunID).ToList();
            List<Test> sortedList = listOfTests.OrderBy(o => o.SuiteName).ThenBy(o => o.CaseID).ToList();
            return sortedList;
        }

        private static List<Case> SortListOfCases(List<Case> listOfCases)
        {
            List<Case> sortedList = listOfCases.OrderByDescending(o => o.CaseID).ToList();
            return sortedList;
        }
	}
}