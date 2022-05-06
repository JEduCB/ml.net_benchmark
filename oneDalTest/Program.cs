using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.ML;
using Microsoft.ML.Data;

using Microsoft.ML.Trainers;
using Microsoft.ML.Transforms;

namespace oneDalTest
{
    using Helpers;
    using Tasks;

    class Program
    {
        static int Main(string[] args)
        {
            bool onedalEnabled = false;
            string mlnetBackend = string.Empty;

            try
            {
                //Parse for valid arguments
                var _args = Arguments.Parse(args, Constants.NumArguments);

                if (_args.Count != Constants.NumArguments)
                {
                    //If parsing failed, show how to use and exit
                    ShowUsageAndExit();
                }

                int iterations = int.Parse(_args[Constants.Iterations]);

                //Save current MLNEM_BACKEND env var value
                mlnetBackend = Environment.GetEnvironmentVariable(Constants.MLNET_BACKEND);

                //Set backend to default
                Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, "");

                //If onedal enabled is passed as argument
                if (_args[Constants.Onedal].Equals(Constants.OnedalEnabled))
                {
                    //Set MLNEM_BACKEND=ONEDAL
                    Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, Constants.ONEDAL);
                    onedalEnabled = true;
                }

                //Run the task based on the task argument
                switch(_args[Constants.Task])
                {
                    case Constants.Binary:
                        Binary.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target], onedalEnabled, iterations);
                        break;

                    case Constants.MultiClass:
                        MultiClass.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target], onedalEnabled, iterations);
                        break;

                    case Constants.Regression:
                        Regression.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target], onedalEnabled, iterations);
                        break;

                    default:
                        throw new Exception("Invalid task name.");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Restore MLNET_BACKEND original value
                Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, mlnetBackend);
            }

            return 0;
        }

        private static void ShowUsageAndExit(int exitCode = -1)
        {
            Console.Clear();

            Console.WriteLine();
            Console.WriteLine($"Usage:");
            Console.WriteLine();
            Console.WriteLine($"oneDalTest.exe {Constants.Task}=<task_name> {Constants.Dataset}=<dataset_name> " +
                $"{Constants.Target}=<column_name> {Constants.Onedal}=<{Constants.OnedalDisabled}|{Constants.OnedalEnabled}> " +
                $"{Constants.Iterations}=<value>");

            Console.WriteLine();
            Console.WriteLine($"Task\t\t=> {Constants.Binary} --> Binary Classification");
            Console.WriteLine($"\t\t   {Constants.MultiClass} --> Multi-class Classification");
            Console.WriteLine($"\t\t   {Constants.Regression} --> Regression");

            Console.WriteLine();
            Console.WriteLine($"{Constants.Dataset}\t\t=> use the dataset name without the '_train' & '_test' suffixes");
            Console.WriteLine($"{Constants.Target}\t\t=> the target column name into the dataset");
            Console.WriteLine($"{Constants.Onedal}\t\t=> {Constants.OnedalEnabled} for oneDAL backend, {Constants.OnedalDisabled} " +
                $"for default backend");
            Console.WriteLine($"{Constants.Iterations}\t=> number of iterations to be run");

            Console.WriteLine();
            Console.WriteLine($"Example\t\t=> oneDalTest.exe {Constants.Task}={Constants.Binary} {Constants.Dataset}=a9a " +
                $"{Constants.Target}=column2 {Constants.Onedal}={Constants.OnedalEnabled} {Constants.Iterations}=3");

            Console.WriteLine();

            Environment.Exit(exitCode);
        }
    }
}
