using System;
using System.Collections.Generic;
using System.IO;

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
            ExcelDocument excelDoc = null;
            StreamWriter csvWriter = null;

            try
            {
                Dictionary<string, string> _args = null;

                try
                {
                    //Parse for valid arguments
                    _args = Arguments.Parse(args, Constants.NumArguments);

                    if (_args.Count != Constants.NumArguments)
                    {
                        throw new Exception("Invalid arguments");
                    }
                }
                catch (Exception)
                {
                    //If parsing failed, show how to use and exit
                    ShowUsageAndExit();
                }

                int iterations = int.Parse(_args[Constants.Iterations]);

                //Save current MLNEM_BACKEND env var value
                mlnetBackend = Environment.GetEnvironmentVariable(Constants.MLNET_BACKEND);

                //Set backend to default
                Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, "");

                int repeatForOneDal = _args[Constants.Onedal].Equals(Constants.OnedalBoth) ? 2 : 1;

                //If onedal enabled is passed as argument
                if (_args[Constants.Onedal].Equals(Constants.OnedalEnabled))
                {
                    //Set MLNEM_BACKEND=ONEDAL
                    Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, Constants.ONEDAL);
                    onedalEnabled = true;
                }

                if (!_args[Constants.ExcelFile].Contains(Constants.NO_FILE))
                {
                    excelDoc = ExcelExporter.CreateExcelDocument(_args[Constants.ExcelFile], _args[Constants.Task],
                        _args[Constants.Dataset]);
                }

                if (!_args[Constants.CsvFile].Contains(Constants.NO_FILE))
                {
                    csvWriter = new StreamWriter(_args[Constants.CsvFile], true);
                }

                for (int i = 0; i < repeatForOneDal; i++)
                {
                    //Run the task based on the task argument
                    switch (_args[Constants.Task])
                    {
                        case Constants.Binary:
                            Binary.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target],
                                onedalEnabled, iterations, csvWriter, excelDoc, i == 1);
                            break;

                        case Constants.MultiClass:
                            MultiClass.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target],
                                onedalEnabled, iterations, csvWriter, excelDoc, i == 1);
                            break;

                        case Constants.Regression:
                            Regression.RunTask(_args[Constants.Dataset], _args[Constants.Task], _args[Constants.Target],
                                onedalEnabled, iterations, csvWriter, excelDoc, i == 1);
                            break;

                        default:
                            throw new Exception("Invalid task name.");
                    }

                    //If run both onedal and default backend is set
                    if (_args[Constants.Onedal].Equals(Constants.OnedalBoth))
                    {
                        //Set MLNEM_BACKEND=ONEDAL
                        Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, Constants.ONEDAL);
                        onedalEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                //Restore MLNET_BACKEND original value
                Environment.SetEnvironmentVariable(Constants.MLNET_BACKEND, mlnetBackend);
            }

            csvWriter?.Close();
            excelDoc?.WorkbookPart.Workbook.Save();
            excelDoc?.Document.Close();

            excelDoc?.Document.Dispose();

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
                $"{Constants.Iterations}=<value> {Constants.CsvFile}=<csv_file_name> {Constants.ExcelFile}=excel_file_name");

            Console.WriteLine();
            Console.WriteLine($"Task\t\t=> {Constants.Binary} --> Binary Classification");
            Console.WriteLine($"\t\t   {Constants.MultiClass} --> Multi-class Classification");
            Console.WriteLine($"\t\t   {Constants.Regression} --> Regression");

            Console.WriteLine();
            Console.WriteLine($"{Constants.Dataset}\t\t=> use the dataset name without the '_train' & '_test' suffixes");
            Console.WriteLine($"{Constants.Target}\t\t=> the target column name into the dataset");
            Console.WriteLine($"{Constants.Onedal}\t\t=> {Constants.OnedalEnabled} for oneDAL backend, {Constants.OnedalDisabled} " +
                $"for default backend. If omitted, both are run");
            Console.WriteLine($"{Constants.Iterations}\t=> number of iterations to be run");
            Console.WriteLine($"{Constants.CsvFile}\t\t=> file name to append the results to");
            Console.WriteLine($"{Constants.ExcelFile}\t\t=> create a new Excel file to export the result to");

            Console.WriteLine();
            Console.WriteLine($"Examples\t=> oneDalTest.exe {Constants.Task}={Constants.Binary} {Constants.Dataset}=a9a " +
                $"{Constants.Target}=column2 {Constants.Onedal}={Constants.OnedalEnabled} {Constants.Iterations}=3");
            Console.WriteLine($"       \t\t=> oneDalTest.exe {Constants.Task}={Constants.Binary} {Constants.Dataset}=a9a " +
                $"{Constants.Target}=column2 {Constants.Onedal}={Constants.OnedalEnabled} {Constants.Iterations}=3 " +
                $"{Constants.CsvFile}=test.csv");

            Console.WriteLine();

            Environment.Exit(exitCode);
        }
    }
}
