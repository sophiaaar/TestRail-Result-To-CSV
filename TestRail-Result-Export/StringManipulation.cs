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

        public static string GetEditorVersion(string rawValue)
        {
            switch (rawValue)
            {
                case "0":
                    return "RC5";
                case "1":
                    return "RC4";
                case "2":
                    return "RC3";
                case "3":
                    return "RC2";
                case "4":
                    return "RC1";
                case "5":
                    return "Beta 1";
                case "6":
                    return "Beta 2";
                case "7":
                    return "Beta 3";
                case "8":
                    return "Beta 4";
                case "9":
                    return "Beta 5";
                case "10":
                    return "Beta 6";
                case "11":
                    return "Beta 7";
                case "12":
                    return "Beta 8";
                case "13":
                    return "Beta 9";
                case "14":
                    return "Beta 10";
                case "15":
                    return "Beta 11";
                case "16":
                    return "Beta 12";
                case "17":
                    return "Beta 13";
                case "18":
                    return "Beta 14";
                case "19":
                    return "Beta 15";
                case "20":
                    return "Alpha 1";
                case "21":
                    return "Alpha 2";
                case "22":
                    return "Alpha 3";
                case "23":
                    return "Alpha 4";
                case "24":
                    return "Alpha 5";
                case "25":
                    return "Alpha 6";
                case "26":
                    return "Alpha 7";
                case "27":
                    return "Alpha 8";
                case "28":
                    return "Alpha 9";
                case "29":
                    return "Alpha 10";
                case "30":
                    return "f1";
                case "31":
                    return "f2";
                case "32":
                    return "f3";
                case "33":
                    return "f4";
                case "34":
                    return "Patch 1";
                case "35":
                    return "Patch 2";
                case "36":
                    return "Patch 3";
                case "37":
                    return "Patch 4";
                case "38":
                    return "Patch 5";
                case "39":
                    return "Patch 6";
                case "40":
                    return "Special Build";
                default:
                    return "Unknown";
            }
        }
    }
}
