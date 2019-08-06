using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trx2Excel.Model
{
    public class UnitTestResult
    {
        public UnitTestResult()
        {
        }

        public UnitTestResult(UnitTestResult copyFrom)
        {
            TestName = copyFrom.TestName;
            Outcome = copyFrom.Outcome;
            Message = copyFrom.Message;
            StrackTrace = copyFrom.StrackTrace;
            NameSpace = copyFrom.NameSpace;
            Owner = copyFrom.Owner;
            AllOwnersString = copyFrom.AllOwnersString;
            FileName = copyFrom.FileName;
        }

        public string TestName { get; set; }
        public string Outcome { get; set; }
        public string Message { get; set; }
        public string StrackTrace { get; set; }
        public string NameSpace { get; set; }
        public string Owner { get; set; }
        public string AllOwnersString { get; set; }
        public string FileName { get; set; }

    }
}
