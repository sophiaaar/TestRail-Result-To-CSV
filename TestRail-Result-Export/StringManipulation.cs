using System;
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
                return "";
            }
        }

        /// <summary>
        /// Converts the status number to a string
        /// </summary>
        /// <returns>The status.</returns>
        /// <param name="rawValue">Raw value.</param>
        public static string GetStatus(string rawValue)
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
    }
}
