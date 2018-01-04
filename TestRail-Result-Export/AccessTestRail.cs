using System;
using Gurock.TestRail;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;

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
    }
}
