using Microsoft.ML;
using Microsoft.ML.Data;
using System;
using System.Collections.Generic;
using System.IO;

namespace oneDalTest.Helpers
{
    internal static class DataLoader
    {
        public static IDataView[] LoadData(MLContext mlContext, string dataset, string task, string label, char separator = ',')
        {
            var trainDataset = $"Data/{dataset}_train.csv";
            var testDataset = $"Data/{dataset}_test.csv";

            if (!File.Exists(trainDataset) || !File.Exists(testDataset))
            {
                Console.WriteLine($"Cannot find '{trainDataset}' or '{testDataset}' dataset.");
                throw new Exception("Dataset not found");
            }

            System.IO.StreamReader file = new System.IO.StreamReader(trainDataset);
            string header = file.ReadLine();
            file.Close();
            
            string[] headerArray = header.Split(separator);
            List<TextLoader.Column> columns = new List<TextLoader.Column>();
            
            foreach (string column in headerArray)
            {
                if (column == label && task != Constants.Regression)
                {
                    if (task == Constants.Binary)
                    {
                        columns.Add(new TextLoader.Column(column, DataKind.Boolean, Array.IndexOf(headerArray, column)));
                    }
                    else
                    {
                        columns.Add(new TextLoader.Column(column, DataKind.UInt32, Array.IndexOf(headerArray, column)));
                    }
                }
                else
                {
                    columns.Add(new TextLoader.Column(column, DataKind.Single, Array.IndexOf(headerArray, column)));
                }
            }

            var loader = mlContext.Data.CreateTextLoader(separatorChar: separator, hasHeader: true, columns: columns.ToArray());

            List<IDataView> dataList = new List<IDataView>();
            dataList.Add(loader.Load(trainDataset));
            dataList.Add(loader.Load(testDataset));

            return dataList.ToArray();
        }

        public static string[] GetFeaturesArray(IDataView data, string labelName)
        {
            List<string> featuresList = new List<string>();

            var nColumns = data.Schema.Count;
            var columnsEnumerator = data.Schema.GetEnumerator();
            
            for (int i = 0; i < nColumns; i++)
            {
                columnsEnumerator.MoveNext();

                if (columnsEnumerator.Current.Name != labelName)
                {
                    featuresList.Add(columnsEnumerator.Current.Name);
                }
            }

            return featuresList.ToArray();
        }
    }
}
