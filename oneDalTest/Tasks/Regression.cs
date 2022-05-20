using System;
using System.Collections.Generic;

namespace oneDalTest.Tasks
{
    using Helpers;
    using Microsoft.ML;
    using System.IO;

    internal class Regression
    {
        static internal void RunTask(string dataset, string task, string targetColumn, bool onedalEnabled, int iterations,
            StreamWriter csvWriter, ExcelDocument excelDoc, bool calculateSpeedUp)
        {
            List<string> rows = null;

            if (excelDoc != null)
            {
                rows = new List<string>();
            }

            string header = "Run,OneDAL,Features,Dataset,Task,All time[ms],Reading time[ms],Fitting time[ms],Prediction time[ms]," +
                "Evaluation time[ms],MAE,RMSE,R2";

            Console.WriteLine();
            Console.WriteLine("Warming up... Please wait!");
            Console.WriteLine();

            for (int i = 0; i <= iterations; ++i)
            {
                var tg = System.Diagnostics.Stopwatch.StartNew();
                var t0 = System.Diagnostics.Stopwatch.StartNew();

                // Create a new context for ML.NET operations. It can be used for
                // exception tracking and logging, as a catalog of available operations
                // and as the source of randomness. Setting the seed to a fixed number
                // in this example to make outputs deterministic.
                var mlContext = new MLContext(seed: 0);

                var data = DataLoader.LoadData(mlContext, dataset, task, targetColumn);
                var featuresArray = DataLoader.GetFeaturesArray(data[0]);

                IDataView trainingData, testingData;

                var preprocessingModel = mlContext.Transforms.Concatenate("Features", featuresArray);
                trainingData = preprocessingModel.Fit(data[0]).Transform(data[0]);
                testingData = preprocessingModel.Fit(data[0]).Transform(data[1]);
                t0.Stop();

                var t1 = System.Diagnostics.Stopwatch.StartNew();
                var trainer = mlContext.Regression.Trainers.Ols(labelColumnName: $"{targetColumn}", featureColumnName: "Features");
                var model = trainer.Fit(trainingData);
                t1.Stop();

                var t2 = System.Diagnostics.Stopwatch.StartNew();
                IDataView predictions = model.Transform(testingData);
                t2.Stop();

                var t3 = System.Diagnostics.Stopwatch.StartNew();
                List<double> metricsList = new List<double>();
                var metrics = mlContext.Regression.Evaluate(predictions, labelColumnName: $"{targetColumn}", scoreColumnName: "Score");
                t3.Stop();
                tg.Stop();

                if (i == 0)
                {
                    ResultConsole.WriteLine($"Run {i} - Warm-up result.", csvWriter, rows);
                    ResultConsole.WriteLine(string.Empty, csvWriter, rows);
                    ResultConsole.WriteLine(header, csvWriter, rows);

                    var result = $"{i},{onedalEnabled},{featuresArray.Length},{dataset},{task}," +
                        $"{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds},{t1.Elapsed.TotalMilliseconds}," +
                        $"{t2.Elapsed.TotalMilliseconds},{t3.Elapsed.TotalMilliseconds},{metrics.MeanAbsoluteError}," +
                        $"{metrics.RootMeanSquaredError},{metrics.RSquared}";

                    ResultConsole.WriteLine(result, csvWriter, rows);

                    Console.WriteLine();
                    Console.WriteLine("Running...");

                    ResultConsole.WriteLine(string.Empty, csvWriter, rows);
                    ResultConsole.WriteLine($"{iterations} test iterations.", csvWriter, rows);

                    ResultConsole.WriteLine(string.Empty, csvWriter, rows);
                    ResultConsole.WriteLine(header, csvWriter, rows);
                }
                else
                {
                    var result = $"{i},{onedalEnabled},{featuresArray.Length},{dataset},{task}," +
                        $"{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds},{t1.Elapsed.TotalMilliseconds}," +
                        $"{t2.Elapsed.TotalMilliseconds},{t3.Elapsed.TotalMilliseconds},{metrics.MeanAbsoluteError}," +
                        $"{metrics.RootMeanSquaredError},{metrics.RSquared}";

                    ResultConsole.WriteLine(result, csvWriter, rows);
                }
            }

            ResultConsole.WriteLine(string.Empty, csvWriter);

            csvWriter?.Flush();

            if (rows != null)
            {
                ExcelExporter.Export(excelDoc, rows, calculateSpeedUp);
            }
        }
    }
}