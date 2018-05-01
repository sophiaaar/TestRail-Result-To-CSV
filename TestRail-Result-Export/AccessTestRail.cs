using System;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Collections.Generic;

namespace TestRailResultExport
{
    public class AccessTestRail
    {
        public static JArray GetRunsForMilestone(APIClient client, string milestoneID)
        {
            return (JArray)client.SendGet("get_runs/2&milestone_id=" + milestoneID);
        }

        public static JArray GetPlansForMilestone(APIClient client, string milestoneID)
        {
            return (JArray)client.SendGet("get_plans/2&milestone_id=" + milestoneID);
        }

        public static JArray GetTestsInRun(APIClient client, string runID)
        {
            return (JArray)client.SendGet("get_tests/" + runID);
        }

        public static JArray GetCasesInSuite(APIClient client, string projectID, string suiteID)
        {
            return (JArray)client.SendGet("get_cases/" + projectID + "&suite_id=" + suiteID);
        }

        public static JArray GetSuitesInProject(APIClient client, string projectID)
        {
            return (JArray)client.SendGet("get_suites/" + projectID);
        }

        public static JArray GetLatestResultsOfTest(APIClient client, string testID, string amountOfResultsToShow)
        {
            return (JArray)client.SendGet("get_results/" + testID + "&limit=" + amountOfResultsToShow);
        }

        public static JArray GetLatestResultsForCase(APIClient client, string runID, string caseID, string amountOfResultsToShow)
        {
            return (JArray)client.SendGet("get_results_for_case/" + runID + "/" + caseID + "&limit=" + amountOfResultsToShow);
        }

        public static JObject GetCase (APIClient client, string caseID)
        {
            return (JObject)client.SendGet($"get_case/" + caseID);
        }

        public static JObject GetSuite(APIClient client, string suiteID)
        {
            return (JObject)client.SendGet($"get_suite/" + suiteID);
        }

        public static JObject GetSection(APIClient client, string sectionID)
        {
            return (JObject)client.SendGet($"get_section/" + sectionID);
        }

        public static JArray GetCaseTypes(APIClient client)
        {
            return (JArray)client.SendGet("get_case_types");
        }
        public static JArray GetStatuses(APIClient client)
        {
            return (JArray)client.SendGet("get_statuses");
        }
        public static JArray GetResultsFields(APIClient client)
        {
            return (JArray)client.SendGet("get_result_fields");
        }


        public static List<int> GetAllSuites(JArray arrayOfSuites)
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

        public static string GetSectionName(JObject section)
        {
            string name = section.Property("name").Value.ToString();
            return name;
        }

        public static List<string> GetRunsInPlan(JArray planArray, APIClient client, List<string> suiteInPlanIDs, List<MainClass.Run> runs)
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

                                MainClass.Run run;
                                run.RunID = runInPlanId;
                                run.Config = runObject.Property("config").Value.ToString();
                                runs.Add(run);

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

        public static void GetSuitesAndRuns(JArray runsArr, List<string> suiteIDs, List<string> runIDs, List<MainClass.Run> runs)
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

                    MainClass.Run run;
                    run.RunID = run_id;
                    run.Config = arrayObject.Property("config").Value.ToString();
                    runs.Add(run);
                }
                else
                {
                    string suite_id = "0";

                    string run_id = arrayObject.Property("id").Value.ToString();

                    suiteIDs.Add(suite_id);
                    runIDs.Add(run_id);

                    MainClass.Run run;
                    run.RunID = run_id;
                    run.Config = arrayObject.Property("config").Value.ToString();
                    runs.Add(run);
                }
            }
        }
    }
}
