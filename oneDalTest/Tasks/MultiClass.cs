﻿using Microsoft.ML;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace oneDalTest.Tasks
{
    using Helpers;
    using Microsoft.ML.Trainers;

    internal class MultiClass
    {
        static internal void RunTask(string dataset, string task, string targetColumn, bool onedalEnabled, int iterations)
        {
            string header = "Run,Dataset,Task,All time[ms],Reading time[ms],Fitting time[ms],Prediction time[ms],Evaluation time[ms]," +
                        "LogLoss,MicroAccuracy,MacroAccuracy";

            Console.WriteLine();
            Console.WriteLine("Running Multi-Class Classification Test");
            Console.WriteLine("Logical CPUs Count = " + Environment.ProcessorCount);
            Console.WriteLine("Using oneDAL = " + onedalEnabled.ToString());
            Console.WriteLine($"Arranging data for task: {task} (configuring preprocessing)");

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
                var metrics = mlContext.MulticlassClassification.Evaluate(predictions, labelColumnName: "target");
                t3.Stop();
                tg.Stop();

                if (i == 0)
                {
                    Console.WriteLine($"Found [{featuresArray.Length}] features.");

                    Console.WriteLine();
                    Console.WriteLine($"First run result - Run {i} (warming up).");

                    Console.WriteLine();
                    Console.WriteLine(header);
                    Console.WriteLine($"{i},{dataset},{task},{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds}," +
                        $"{t1.Elapsed.TotalMilliseconds},{t2.Elapsed.TotalMilliseconds},{t3.Elapsed.TotalMilliseconds}," +
                        $"{metrics.LogLoss},{metrics.MicroAccuracy},{metrics.MacroAccuracy}");

                    Console.WriteLine();
                    Console.WriteLine($"{iterations} iterations will be now executed.");

                    Console.WriteLine();
                    Console.WriteLine(header);
                }
                else
                {
                    Console.Write($"{i},{dataset},{task},{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds}," +
                        $"{t1.Elapsed.TotalMilliseconds},{t2.Elapsed.TotalMilliseconds},{t3.Elapsed.TotalMilliseconds}," +
                        $"{metrics.LogLoss},{metrics.MicroAccuracy},{metrics.MacroAccuracy}\n");
                }
            }
        }
    }
}