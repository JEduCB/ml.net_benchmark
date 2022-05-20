using Microsoft.ML;
using System;
using System.Collections.Generic;
using System.Linq;

namespace oneDalTest.Tasks
{
    using Helpers;
    using Microsoft.ML.Trainers;
    using System.IO;

    internal class MultiClass
    {
        static internal void RunTask(string dataset, string task, string targetColumn, bool onedalEnabled, int iterations,
            StreamWriter csvWriter, ExcelDocument excelDoc, bool calculateSpeedUp)
        {
            List<string> rows = null;

            if (excelDoc != null)
            {
                rows = new List<string>();
            }

            string header = "Run,OneDAL,Features,Dataset,Task,All time[ms],Reading time[ms],Fitting time[ms]," +
                "Prediction time[ms],Evaluation time[ms],LogLoss,MicroAccuracy,MacroAccuracy";

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
                var featuresArray = DataLoader.GetFeaturesArray(data[0], targetColumn);

                IDataView trainingData, testingData;

                var preprocessingModel = mlContext.Transforms.Concatenate("Features",
                    featuresArray).Append(mlContext.Transforms.Conversion.MapValueToKey(targetColumn));

                trainingData = preprocessingModel.Fit(data[0]).Transform(data[0]);
                testingData = preprocessingModel.Fit(data[0]).Transform(data[1]);
                t0.Stop();

                var t1 = System.Diagnostics.Stopwatch.StartNew();
                var options = new LbfgsMaximumEntropyMulticlassTrainer.Options()
                {
                    LabelColumnName = targetColumn,
                    FeatureColumnName = "Features",
                    L1Regularization = 0.0f,
                    L2Regularization = 0.0f,
                    HistorySize = 1,
                    OptimizationTolerance = 1e-12f,
                    MaximumNumberOfIterations = 100,
                    NumberOfThreads = Environment.ProcessorCount
                };

                var trainer = mlContext.MulticlassClassification.Trainers.LbfgsMaximumEntropy(options);
                var model = trainer.Fit(trainingData);
                t1.Stop();

                var t2 = System.Diagnostics.Stopwatch.StartNew();
                IDataView predictions = model.Transform(testingData);
                t2.Stop();

                var t3 = System.Diagnostics.Stopwatch.StartNew();
                var metrics = mlContext.MulticlassClassification.Evaluate(predictions, labelColumnName: $"{targetColumn}");
                t3.Stop();
                tg.Stop();

                if (i == 0)
                {
                    ResultConsole.WriteLine($"Run {i} - Warm-up result.", csvWriter, rows);
                    ResultConsole.WriteLine(string.Empty, csvWriter, rows);
                    ResultConsole.WriteLine(header, csvWriter, rows);

                    var result = $"{i},{onedalEnabled},{featuresArray.Length},{dataset},{task}," +
                        $"{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds}," +
                        $"{t1.Elapsed.TotalMilliseconds},{t2.Elapsed.TotalMilliseconds}," +
                        $"{t3.Elapsed.TotalMilliseconds},{metrics.LogLoss},{metrics.MicroAccuracy}," +
                        $"{metrics.MacroAccuracy}";

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
                        $"{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds}," +
                        $"{t1.Elapsed.TotalMilliseconds},{t2.Elapsed.TotalMilliseconds}," +
                        $"{t3.Elapsed.TotalMilliseconds},{metrics.LogLoss},{metrics.MicroAccuracy}," +
                        $"{metrics.MacroAccuracy}";

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