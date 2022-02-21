using Microsoft.ML;
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
            Console.WriteLine();

            //Temporary read just for getting the number of features - Start
            var tmpContext = new MLContext(seed: 0);
            var tmpData = DataLoader.LoadData(tmpContext, dataset, task, targetColumn);
            var tmpFeaturesArray = DataLoader.GetFeaturesArray(tmpData[0]);
            //Temporary read just for getting the number of features - End

            Console.WriteLine();
            Console.WriteLine("Running Multi-Class Classification Test");
            Console.WriteLine("Using oneDAL = " + onedalEnabled.ToString());
            Console.WriteLine($"Found [{tmpFeaturesArray.Length}] features.");
            Console.WriteLine($"Arranging data for task: {task} (configuring preprocessing).");

            Console.WriteLine();
            Console.WriteLine("Dataset,Task,All time[ms],Reading time[ms],Fitting time[ms],Prediction time[ms],Evaluation time[ms]," +
                "LogLoss,MicroAccuracy,MacroAccuracy");

            for (int i = 0; i < iterations; ++i)
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
                    L1Regularization = 0.05f,
                    L2Regularization = 0.05f,
                    HistorySize = 20,
                    OptimizationTolerance = 1e-6f,
                    MaximumNumberOfIterations = 100
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

                Console.Write($"{dataset},{task},{tg.Elapsed.TotalMilliseconds},{t0.Elapsed.TotalMilliseconds}," +
                    $"{t1.Elapsed.TotalMilliseconds},{t2.Elapsed.TotalMilliseconds},{t3.Elapsed.TotalMilliseconds}," +
                    $"{metrics.LogLoss},{metrics.MicroAccuracy},{metrics.MacroAccuracy}\n");
            }
        }
    }
}