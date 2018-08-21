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
        public static JArray GetRunsForMilestone(APIClient client, string projectID, string milestoneID)
        {
            return (JArray)client.SendGet("get_runs/" + projectID + "&milestone_id=" + milestoneID);
        }
        
		public static JArray GetRuns(APIClient client, string projectID)
        {
            return (JArray)client.SendGet("get_runs/" + projectID);
        }

        public static JArray GetPlansForMilestone(APIClient client, string projectID, string milestoneID)
        {
            return (JArray)client.SendGet("get_plans/" + projectID + "&milestone_id=" + milestoneID);
        }
        
		public static JArray GetPlans(APIClient client, string projectID)
        {
            return (JArray)client.SendGet("get_plans/" + projectID);
        }

		public static JArray GetMilestones(APIClient client, string projectID)
        {
            return (JArray)client.SendGet("get_milestones/" + projectID);
        }

		public static JObject GetMilestone(APIClient client, string milestoneID)
        {
			return (JObject)client.SendGet("get_milestone/" + milestoneID);
        }

		public static JObject GetProject(APIClient client, string projectID)
        {
			return (JObject)client.SendGet("get_project/" + projectID);
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

        public static JObject GetCase(APIClient client, string caseID)
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

        public static List<string> GetRunsInPlan(JArray planArray, APIClient client, List<string> suiteInPlanIDs, List<string> runNames, List<MainClass.Run> runs)
        {
            List<JArray> ListOfRunsInPlan = new List<JArray>();
            List<string> planIds = new List<string>();
            List<string> runInPlanIds = new List<string>();

            for (int i = 0; i < planArray.Count; i++)
            {
                JObject arrayObject = planArray[i].ToObject<JObject>();

                string planID = arrayObject.Property("id").Value.ToString();
                planIds.Add(planID);

                string planName = arrayObject.Property("name").Value.ToString();

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
                                run.isCompleted = runObject.Property("is_completed").Value.ToString();
								run.MilestoneID = runObject.Property("milestone_id").Value.ToString();
                                runs.Add(run);

                                if (!runInPlanIds.Contains(runInPlanId))
                                {
                                    runInPlanIds.Add(runInPlanId);

                                    //string runName = runObject.Property("name").Value.ToString();
                                    runNames.Add(planName);

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

        public static void GetSuitesAndRuns(JArray runsArr, List<string> suiteIDs, List<string> runIDs, List<string> runNames, List<MainClass.Run> runs)
        {
            for (int i = 0; i < runsArr.Count; i++)
            {
                JObject arrayObject = runsArr[i].ToObject<JObject>();

                JProperty prop = arrayObject.Property("suite_id");
                if (prop != null && prop.Value != null && !string.IsNullOrEmpty(prop.Value.ToString()))
                {
                    string suite_id = arrayObject.Property("suite_id").Value.ToString();

                    string run_id = arrayObject.Property("id").Value.ToString();
                    string run_name = arrayObject.Property("name").Value.ToString();

                    suiteIDs.Add(suite_id);
                    runIDs.Add(run_id);
                    runNames.Add(run_name);

                    MainClass.Run run;
                    run.RunID = run_id;
                    run.Config = arrayObject.Property("config").Value.ToString();
                    run.isCompleted = arrayObject.Property("is_completed").Value.ToString();
					run.MilestoneID = arrayObject.Property("milestone_id").Value.ToString();
                    runs.Add(run);
                }
                else
                {
                    string suite_id = "0";

                    string run_id = arrayObject.Property("id").Value.ToString();
                    string run_name = arrayObject.Property("name").Value.ToString();

                    suiteIDs.Add(suite_id);
                    runIDs.Add(run_id);
                    runNames.Add(run_name);

                    MainClass.Run run;
                    run.RunID = run_id;
                    run.Config = arrayObject.Property("config").Value.ToString();
                    run.isCompleted = arrayObject.Property("is_completed").Value.ToString();
					run.MilestoneID = arrayObject.Property("milestone_id").Value.ToString();
                    runs.Add(run);
                }
            }
        }

		public static string GetEditorVersion(APIClient client, string projectID, string rawValue)
        {
            JArray resultFieldsArray = GetResultsFields(client);

			for (int i = 0; i < resultFieldsArray.Count; i++)
			{
				JObject resultObject = resultFieldsArray[i].ToObject<JObject>();

				if (resultObject.Property("name").Value.ToString() == "editorversion")
				{
					JProperty configs = resultObject.Property("configs");
                    
					foreach (JArray child in configs.OfType<JArray>())
					//for (int k = 0; k < configs.OfType<JArray>().Count<JArray>(); k++)
					{
						//JObject context = (JObject)child["context"];
						//JObject contextOuter = (JObject)child[0];
						for (int k = 0; k < child.Count; k++)
						{
							JObject contextInner = (JObject)child[k];
                            
							for (int m = 0; m < contextInner.Count; m++)
							{
								JObject context = (JObject)contextInner["context"];
								JArray projectIds = (JArray)context["project_ids"];


								for (int j = 0; j < projectIds.Count; j++)
								{
									var projectObject = projectIds[j];
									if (projectObject.ToString() == projectID)
									{
										//get list of editor versions
										JObject options = (JObject)contextInner["options"];

										string versions = options.Property("items").Value.ToString();
                                        string[] editorVersions = versions.ToString().Split('\n');
                                        foreach (string editorVersion in editorVersions)
                                        {
                                            string[] values = editorVersion.Split(',');
                                            string id = values[0];
                                            string name = values[1];
                                            if (id == rawValue)
                                            {
                                                return name;
                                            }
                                        }

										//for (int n = 0; n < options.Count; n++)
										//{

										//	JArray versions = (JArray)options["items"];
										//	string[] editorVersions = versions.ToString().Split('\n');
										//	foreach (string editorVersion in editorVersions)
										//	{
										//		string[] values = editorVersion.Split(',');
										//		string id = values[0];
										//		string name = values[1];
										//		if (id == rawValue)
										//		{
										//			return name;
										//		}
										//	}
										//}
									}
								}
							}
						}

					}


				}
			}
			return "";
        }
    }
}
