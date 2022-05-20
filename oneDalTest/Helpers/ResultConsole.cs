using System;
using System.Collections.Generic;
using System.IO;

namespace oneDalTest.Helpers
{
    internal static class ResultConsole
    {
        public static void WriteLine(string value, StreamWriter csvWriter, List<string> xlsxRows = null)
        {
            Console.WriteLine(value);
            csvWriter?.WriteLine(value);
            xlsxRows?.Add(value);
        }
    }
}
