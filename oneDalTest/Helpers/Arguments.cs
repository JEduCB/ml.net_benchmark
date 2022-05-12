using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace oneDalTest.Helpers
{   
    internal class Arguments
    {
        public static Dictionary<string, string> Parse(string [] args, int numArgs)
        {
            string[] _argumentList = { Constants.Task, Constants.Dataset, Constants.Target, Constants.Onedal, Constants.Iterations };

            //If an optional argumen is not explicitely set, set it to its default value
            var argsList = args.ToList();
            if(argsList.Count(arg => arg.Contains(Constants.Iterations)) == 0)
            {
                argsList.Add(Constants.Iterations + "=1");
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
            string[] _onedalValues = { Constants.OnedalDisabled, Constants.OnedalEnabled };

            return keyValuePairArg[0] == Constants.Task && _taskValues.Contains(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Dataset && !string.IsNullOrEmpty(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Target && !string.IsNullOrEmpty(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Onedal && _onedalValues.Contains(keyValuePairArg[1])
                   || keyValuePairArg[0] == Constants.Iterations && !string.IsNullOrEmpty(keyValuePairArg[1]) &&
                        uint.TryParse(keyValuePairArg[1], out uint r) && r > 0;
                   //add more validations here
                   //|| pairValueArg[0] == "<new_arg_to_validade>" && _<new_valid_arg_values_list>.Contains(pairValueArg[1]);
        }
    }
}
