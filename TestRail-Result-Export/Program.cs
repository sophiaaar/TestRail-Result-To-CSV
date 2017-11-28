using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TestRailResultExport
{
    class MainClass
    {
        public static string milestoneID = "";
        public static List<int> suiteIDs = new List<int>();
        public static List<int> runIDs = new List<int>();

        public static List<int> suiteInPlanIDs = new List<int>();

        private static readonly IConfigReader _configReader = new ConfigReader();

        //private static CsvBuilders _csvBuilder;

        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            APIClient client = ConnectToTestrail();

            GetAllTests(client);
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
            //APIClient client = ConnectToTestrail();
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();
            //JArray c = (JArray)client.SendGet("get_runs/2&milestone_id=" + milestoneID);
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

			for (int i = 0; i < runIDs.Count; i++)
			{
                //JArray testsArray = (JArray)client.SendGet("get_tests/" + runIDs[i]);
                JArray testsArray = GetTestsInRun(client, runIDs[i].ToString());

                string suiteName = "";

                if (suiteIDs[i] != 0)
                {
                    JObject suite = (JObject)client.SendGet($"get_suite/{suiteIDs[i]}");
                    suiteName = suite.Property("name").Value.ToString();
                }
                else
                {
                    suiteName = "deleted";
                }

				string csvOfTests = CreateCSVOfTests(testsArray, suiteIDs[i], suiteName);
				Console.WriteLine(csvOfTests);

			}

            List<string> runInPlanIds = GetRunsInPlan(planArray, client);

			for (int i = 0; i < runInPlanIds.Count; i++)
			{
                //JArray testsArray = (JArray)client.SendGet("get_tests/" + runInPlanIds[i]);
                JArray testsArray = GetTestsInRun(client, runInPlanIds[i].ToString());

                string suiteName = "";

                if (suiteInPlanIDs[i] != 0)
                {
                    JObject suite = (JObject)client.SendGet($"get_suite/{suiteInPlanIDs[i]}");
                    suiteName = suite.Property("name").Value.ToString();
                }
                else
                {
                    suiteName = "deleted";
                }

				string csvOfTests = CreateCSVOfTests(testsArray, suiteInPlanIDs[i], suiteName);
				Console.WriteLine(csvOfTests);

			}

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
                                    suiteInPlanIDs.Add(Int32.Parse(suiteInPlanId));
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
                    int suite_id = Int32.Parse(arrayObject.Property("suite_id").Value.ToString());

                    int run_id = Int32.Parse(arrayObject.Property("id").Value.ToString());

                    suiteIDs.Add(suite_id);
                    runIDs.Add(run_id);
                }
                else
                {
					int suite_id = 0;

					int run_id = Int32.Parse(arrayObject.Property("id").Value.ToString());

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
			//csv.Append(header);

			for (int i = 0; i < rawData.Count; i++)
			{
				JObject arrayObject = rawData[i].ToObject<JObject>();

				string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", arrayObject.Property("plan_id").Value, arrayObject.Property("suite_id").Value, arrayObject.Property("name").Value, arrayObject.Property("is_completed").Value, arrayObject.Property("passed_count").Value, arrayObject.Property("failed_count").Value, arrayObject.Property("blocked_count").Value, arrayObject.Property("custom_status1_count").Value, arrayObject.Property("untested_count").Value, arrayObject.Property("milestone_id").Value, arrayObject.Property("url").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}


		public static string CreateCSVOfTests(JArray arrayOfTests, int suiteId, string suiteName)
		{
			StringBuilder csv = new StringBuilder();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Status", "Estimate","\n");
			csv.Append(header);

			for (int i = 0; i < arrayOfTests.Count; i++)
			{
				JObject arrayObject = arrayOfTests[i].ToObject<JObject>();
                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString() + "\"", GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value, "\n");
				csv.Append(newLine);
			}

			return csv.ToString();
		}
    }
}
