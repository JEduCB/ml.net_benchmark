using System;
using System.Collections.Generic;
using System.Linq;

namespace oneDalTest.Helpers
{
    internal class Arguments
    {
        public static Dictionary<string, string> Parse(string [] args, int numArgs)
        {
            string[] _argumentList = { Constants.Task, Constants.Dataset, Constants.Target, Constants.Onedal,
                Constants.Iterations, Constants.CsvFile, Constants.ExcelFile };

            //If an optional argumen is not explicitely set, set it to its default value
            var argsList = args.ToList();
            if (argsList.Count(arg => arg.Contains(Constants.Onedal)) == 0)
            {
                argsList.Add(Constants.Onedal + "=2");
            }
            if (argsList.Count(arg => arg.Contains(Constants.Iterations)) == 0)
            {
                argsList.Add(Constants.Iterations + "=1");
            }
            if (argsList.Count(arg => arg.Contains(Constants.CsvFile)) == 0)
            {
                argsList.Add(Constants.CsvFile + "=" + Constants.NO_FILE);
            }
            if (argsList.Count(arg => arg.Contains(Constants.ExcelFile)) == 0)
            {
                argsList.Add(Constants.ExcelFile + "=" + Constants.NO_FILE);
            }
            args = argsList.ToArray();
            //

            var arguments = new Dictionary<string, string>();

            if(args.Length == numArgs)
            {
                foreach(var arg in args)
                {
                    var keyValuePairArg = arg.Split('=');
                    keyValuePairArg[0] = keyValuePairArg[0].ToLower();

                    if (_argumentList.Contains(keyValuePairArg[0]) && IsValidPairValue(keyValuePairArg))
                    {
                        if (!arguments.ContainsKey(keyValuePairArg[0]))
                        {
                            arguments.Add(keyValuePairArg[0], keyValuePairArg[1]);
                        }
                    }
                }
            }

            return arguments;
        }

        private static bool IsValidPairValue(string[] keyValuePairArg)
        {
            string[] _taskValues = { Constants.MultiClass, Constants.Regression, Constants.Binary };
            string[] _onedalValues = { Constants.OnedalDisabled, Constants.OnedalEnabled, Constants.OnedalBoth };

            return keyValuePairArg[0] == Constants.Task && _taskValues.Contains(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Dataset && !string.IsNullOrEmpty(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Target && !string.IsNullOrEmpty(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Onedal && _onedalValues.Contains(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Iterations && !string.IsNullOrEmpty(keyValuePairArg[1]) &&
                        uint.TryParse(keyValuePairArg[1], out uint r) && r > 0
                   || keyValuePairArg[0] == Constants.CsvFile && !string.IsNullOrEmpty(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.ExcelFile && !string.IsNullOrEmpty(keyValuePairArg[1]);
            //add more validations here
            //|| keyValuePairArg[0] == "<new_arg_to_validade>" && _condition);
        }
    }
}
