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
                return "Valid";
            }
        }

        /// <summary>
        /// Converts the status number to a string
        /// </summary>
        /// <returns>The status.</returns>
        /// <param name="rawValue">Raw value.</param>
        public static string GetStatus(string rawValue)
        {
            switch (rawValue)
            {
                case "1":
                    return "Passed";
                case "2":
                    return "Blocked";
                case "3":
                    return "Untested";
                case "4":
                    return "Retest";
                case "5":
                    return "Failed";
                case "6":
                    return "Pending";
                default:
                    return "Other";
            }
        }

        public static string GetCaseType(string rawValue)
        {
            switch (rawValue)
            {
                case "1":
                    return "Acceptance";
                case "2":
                    return "Accessibility";
                case "3":
                    return "Automated";
                case "4":
                    return "Compatibility";
                case "5":
                    return "Destructive";
                case "6":
                    return "Functional";
                case "7":
                    return "Other";
                case "8":
                    return "Performance";
                case "9":
                    return "Regression";
                case "10":
                    return "Security";
                case "11":
                    return "Smoke and Sanity";
                case "12":
                    return "Usability";
                case "13":
                    return "Use Case";
                default:
                    return "Unknown";
            }
        }
    }
}
