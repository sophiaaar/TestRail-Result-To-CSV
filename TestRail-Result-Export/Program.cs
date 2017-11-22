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

        private static readonly IConfigReader _configReader = new ConfigReader();

        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            APIClient client = ConnectToTestrail();

            //milestone 88 is 2017.3
            JArray arrayOfSuites = (JArray)client.SendGet("get_suites/2");

            List<int> suiteIds = GetAllSuites(arrayOfSuites);
            GetAllTests();



        }

        private static APIClient ConnectToTestrail()
        {
			APIClient client = new APIClient("http://qatestrail.hq.unity3d.com");
			client.User = _configReader.TestRailUser;
			client.Password = _configReader.TestRailPass;
            return client;
        }

        private static void GetAllTests()
        {
            APIClient client = ConnectToTestrail();
			Console.WriteLine("Enter milestone ID: ");
			milestoneID = Console.ReadLine();
			JArray c = (JArray)client.SendGet("get_runs/2&milestone_id=" + milestoneID);


			FileStream ostrm;
			StreamWriter writer;
			TextWriter oldOut = Console.Out;
			try
			{
				ostrm = new FileStream("ResultsOverview.csv", FileMode.OpenOrCreate, FileAccess.Write);
				writer = new StreamWriter(ostrm);
			}
			catch (Exception e)
			{
				Console.WriteLine("Cannot open Results.csv for writing");
				Console.WriteLine(e.Message);
				return;
			}
			Console.SetOut(writer);

			string rawCsv = CreateCSVOfRuns(c);
			Console.WriteLine(rawCsv);

			Console.SetOut(oldOut);
			writer.Close();
			ostrm.Close();

			FileStream ostrm2;
			StreamWriter writer2;
			TextWriter oldOut2 = Console.Out;

			try
			{
				ostrm2 = new FileStream("Tests.csv", FileMode.OpenOrCreate, FileAccess.Write);
				writer2 = new StreamWriter(ostrm2);
			}
			catch (Exception e)
			{
				Console.WriteLine("Cannot open Results.csv for writing");
				Console.WriteLine(e.Message);
				return;
			}
			Console.SetOut(writer2);

			//int fileNum = 1;
			// (int runID in runIDs)
			for (int i = 0; i < runIDs.Count; i++)
			{
				JArray testsArray = (JArray)client.SendGet("get_tests/" + runIDs[i]);




				JObject suite = (JObject)client.SendGet($"get_suite/{suiteIDs[i]}");
				string suiteName = suite.Property("name").Value.ToString();

				string csvOfTests = CreateCSVOfTests(testsArray, suiteIDs[i], suiteName);
				Console.WriteLine(csvOfTests);


				//fileNum++;

			}

			Console.SetOut(oldOut2);
			writer2.Close();
			ostrm2.Close();
			Console.WriteLine("Done");
        }

        private static string CreateCSVOfRuns(JArray rawData)
        {
            StringBuilder csv = new StringBuilder();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", "Run ID", "Suite ID", "Title", "Is_Completed", "Passed_count", "Failed_count", "Blocked_count", "pending_count", "Untested_count", "Milestone ID", "url", "\n");
            csv.Append(header);

            JArray parsedArray = ParseRuns(rawData);
            for (int i = 0; i < parsedArray.Count; i++)
            {
                JObject arrayObject = parsedArray[i].ToObject<JObject>();
                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}", arrayObject.Property("id").Value, arrayObject.Property("suite_id").Value, arrayObject.Property("name").Value, arrayObject.Property("is_completed").Value, arrayObject.Property("passed_count").Value, arrayObject.Property("failed_count").Value, arrayObject.Property("blocked_count").Value, arrayObject.Property("custom_status1_count").Value, arrayObject.Property("untested_count").Value, arrayObject.Property("milestone_id").Value, arrayObject.Property("url").Value, "\n");
                csv.Append(newLine);
            }

            return csv.ToString();
        }

        private static string CreateCSVOfTests(JArray arrayOfTests, int suiteId, string suiteName)
        {
            StringBuilder csv = new StringBuilder();

            string header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Suite ID", "Suite Name", "Run ID", "Test ID", "Case ID", "Title", "Status", "Estimate", "\n");
			csv.Append(header);

            for (int i = 0; i < arrayOfTests.Count; i++)
            {
                JObject arrayObject = arrayOfTests[i].ToObject<JObject>();
                string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", suiteId, suiteName, arrayObject.Property("run_id").Value, arrayObject.Property("id").Value, arrayObject.Property("case_id").Value, "\"" + arrayObject.Property("title").Value.ToString()+ "\"", GetStatus(arrayObject.Property("status_id").Value.ToString()), arrayObject.Property("estimate").Value, "\n");
                csv.Append(newLine);
            }

            return csv.ToString();
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

		public static JObject ParseRun(JObject jObj)
		{
			jObj.Property("assignedto_id").Remove();
			jObj.Property("config").Remove();
			jObj.Property("config_ids").Remove();
			jObj.Property("include_all").Remove();
			jObj.Property("custom_status2_count").Remove();
			jObj.Property("custom_status3_count").Remove();
			jObj.Property("custom_status4_count").Remove();
			jObj.Property("custom_status5_count").Remove();
			jObj.Property("custom_status6_count").Remove();
			jObj.Property("custom_status7_count").Remove();
			jObj.Property("created_on").Remove();
			jObj.Property("created_by").Remove();
			return jObj;
		}

		public static JArray ParseRuns(JArray array)
		{
			for (int i = 0; i < array.Count; i++)
			{
				JObject arrayObject = array[i].ToObject<JObject>();

                int suite_id = Int32.Parse(arrayObject.Property("suite_id").Value.ToString());

                int run_id = Int32.Parse(arrayObject.Property("id").Value.ToString());

                suiteIDs.Add(suite_id);
                runIDs.Add(run_id);

    //            if (!suiteIDs.Contains(suite_id))
    //            {
    //                suiteIDs.Add(suite_id);
    //            }

    //            if (!runIDs.Contains(run_id))
				//{
    //                runIDs.Add(run_id);
				//}

				arrayObject = ParseRun(arrayObject);
			}
			return array;
		}
    }
}
