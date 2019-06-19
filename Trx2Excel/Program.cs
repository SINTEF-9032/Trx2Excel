using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net.Repository.Hierarchy;
using Trx2Excel.ExcelUtils;
using Trx2Excel.TrxReaderUtil;
using Trx2Excel.Model;

namespace Trx2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length != 2)
               throw new Exception("Illegal Number of Argument");

            var resultList = new SortedDictionary<string, SortedDictionary<string, UnitTestResult>>();
            var trx_file_name_list = Directory.GetFiles(args[0], "*.trx");
            int pass_count = 0;
            int fail_count = 0;
            int skip_count = 0;
            foreach (string file_name in trx_file_name_list)
            {
                var reader = new TrxReader(file_name);
                Console.WriteLine("[INFO] : Reading the Trx file : {0}", file_name);
                reader.GetTestResults(ref resultList);
                Console.WriteLine("[INFO] : Getting TestResult from Trx file : {0}", file_name);
                pass_count += reader.PassCount;
                fail_count += reader.FailCount;
                skip_count += reader.SkipCount;
            }
            var excelWriter = new ExcelWriter(args[1]);
            excelWriter.WriteToExcel(resultList);
            Console.WriteLine("[INFO] : Writing to Excel File : {0}", args[1]);
            excelWriter.AddChart(pass_count, fail_count, skip_count);
            Console.WriteLine("[INFO] : Generating charts : {0}", args[1]);
            Console.WriteLine("[INFO] : Output File : {0}", args[1]);
        }
    }
}
