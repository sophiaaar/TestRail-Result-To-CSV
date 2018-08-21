using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace TestRailResultExport
{
    public class StringManipulation
    {
        public static string HasSteps(JObject arrayObject)
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

        public static string HasStepsSeparated(JObject arrayObject)
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

        public static string IsInvalid(JObject arrayObject)
        {
            if (HasSteps(arrayObject) == "No" && HasStepsSeparated(arrayObject) == "No")
            {
                return "Invalid";
            }
            else
            {
                return "Valid";
            }
        }

        public static string GetStatus(JArray statusArray, string rawValue)
        {
            string statusName = "";

            for (int i = 0; i < statusArray.Count; i++)
            {
                JObject caseType = statusArray[i].ToObject<JObject>();

                if (caseType.Property("id").Value.ToString() == rawValue)
                {
                    statusName = caseType.Property("name").Value.ToString();

                    if (statusName == "untested")
                    {
                        statusName = "In Progress";
                    }
                    break;
                }
            }

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

            statusName = textInfo.ToTitleCase(statusName);

            return statusName;
        }

        /// <summary>
        /// Converts the status number to a string
        /// </summary>
        /// <returns>The status.</returns>
        /// <param name="rawValue">Raw value.</param>
        public static string GetTemplateStatus(string rawValue)
        {
            switch (rawValue)
            {
                case "1":
                    return "Draft";
                case "2":
                    return "Active";
                case "3":
                    return "Needs Improvement";
                case "4":
                    return "Sustained Engineering";
                case "5":
                    return "Retired";
                default:
                    return "Other";
            }
        }

        public static string GetCaseType(JArray caseTypesArray, string rawValue)
        {
            string caseTypeName = "";

            for (int i = 0; i < caseTypesArray.Count; i++)
            {
                JObject caseType = caseTypesArray[i].ToObject<JObject>();

                if (caseType.Property("id").Value.ToString() == rawValue)
                {
                    caseTypeName = caseType.Property("name").Value.ToString();
                    break;
                }
            }

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

            caseTypeName = textInfo.ToTitleCase(caseTypeName);

            return caseTypeName;
        }
        

        public static int ConvertTimespanStringToSeconds(string timespanString)
		{
			string[] segments = timespanString.Split(' ');

			int dayInSeconds = 86400;
			int hourInSeconds = 3600;
			int minuteInSeconds = 60;

			int totalSeconds = 0;

			foreach (string segment in segments)
			{
				if (segment.Contains("d"))
				{
					string number = segment.TrimEnd('d');
					int days = Int32.Parse(number);
					int daysInSeconds = days * dayInSeconds;
					totalSeconds += daysInSeconds;
				}
				else if (segment.Contains("h"))
				{
                    string number = segment.TrimEnd('h');
					int hours = Int32.Parse(number);
					int hoursInSeconds = hours * hourInSeconds;
					totalSeconds += hoursInSeconds;
				}
				else if (segment.Contains("m"))
				{
                    string number = segment.TrimEnd('m');
					int minutes = Int32.Parse(number);
					int minutesInSeconds = minutes * minuteInSeconds;
					totalSeconds += minutesInSeconds;
				}
				else if (segment.Contains("s"))
				{
                    string number = segment.TrimEnd('s');
					int seconds = Int32.Parse(number);

					totalSeconds += seconds;
				}
			}

			//TimeSpan timeSpan = TimeSpan.Parse(timespanString);
			//double seconds = timeSpan.TotalSeconds;

			return totalSeconds;
		}
    }
}
