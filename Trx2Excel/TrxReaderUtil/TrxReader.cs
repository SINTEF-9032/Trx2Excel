using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Trx2Excel.Model;
using Trx2Excel.Setting;

namespace Trx2Excel.TrxReaderUtil
{
    public class TrxReader
    {
        private string FileName { get; set; }
        public int PassCount { get; set; }
        public int FailCount { get; set; }
        public int SkipCount { get; set; }
        public TrxReader(string fileName)
        {
            FileName = fileName;
        }

        public SortedDictionary<string, SortedDictionary<string, UnitTestResult>> GetTestResults(ref SortedDictionary<string, SortedDictionary<string, UnitTestResult>> total_lists)
        {
            //var resultList = new List<UnitTestResult>();
            var doc = new XmlDocument();
            doc.Load(FileName);
            var xmlNodeList = doc.GetElementsByTagName(NodeName.UnitTestResult);

            if (xmlNodeList.Count <= 0)
                return total_lists;

            foreach (XmlNode node in xmlNodeList)
            {
                var loc_results = GetResults(doc, node);
                foreach (UnitTestResult result in loc_results)
                {
                    result.FileName = FileName;
                    if (!total_lists.ContainsKey(result.Owner))
                    {
                        total_lists.Add(result.Owner, new SortedDictionary<string, UnitTestResult>());
                    }
                    string unique_test_name = result.NameSpace + "." + result.TestName;
                    total_lists[result.Owner].Add(unique_test_name, result);
                }
            }
            return total_lists;
        }

        public List<UnitTestResult> GetResults(XmlDocument doc,XmlNode node)
        {
            var ret_list = new List<UnitTestResult>();

            var result = new UnitTestResult();
            result.TestName = node.Attributes?[NodeName.TestName]?.InnerText;
            result.Outcome = node.Attributes?[NodeName.Outcome]?.InnerText;
            var outcome = (TestOutcome)Enum.Parse(typeof(TestOutcome), result.Outcome, true);
            result.NameSpace = GetNameSpace(doc.GetElementsByTagName(NodeName.UnitTest), node.Attributes?[NodeName.TestId]?.InnerText);
            switch (outcome)
            {
                case TestOutcome.Failed:
                    var output = node.ChildNodes[GetNodeIndex(node, NodeName.Output)];
                    var errorInfo = output.ChildNodes[GetNodeIndex(output, NodeName.ErrorInfo)];
                    result.Message = errorInfo.ChildNodes[GetNodeIndex(errorInfo, NodeName.Message)]?.InnerText;
                    result.StrackTrace = errorInfo.ChildNodes[GetNodeIndex(node, NodeName.StackTrace)]?.InnerText;
                    FailCount++;
                    break;
                case TestOutcome.Passed:
                    PassCount++;
                    break;
                case TestOutcome.Skipped:
                    SkipCount++;
                    break;
            }

            var owners = GetOwners(doc.GetElementsByTagName(NodeName.UnitTest), node.Attributes?[NodeName.TestId]?.InnerText);
            foreach (string owner in owners)
            {
                if (result.AllOwnersString == null)
                {
                    result.AllOwnersString = "";
                }

                if (result.AllOwnersString.Length != 0)
                {
                    result.AllOwnersString += " | ";
                }
                result.AllOwnersString += owner;
            }

            foreach (string owner in owners)
            {
                var entry = new UnitTestResult(result);
                entry.Owner = owner;
                ret_list.Add(entry);
            }
            return ret_list;
        }

        public int GetNodeIndex(XmlNode node, string nodeName)
        {
            for (var i = 0; i < node.ChildNodes.Count; i++)
            {
                if (node.ChildNodes[i].Name.Equals(nodeName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return 0;
        }

        public string GetNameSpace(XmlNodeList list, string id)
        {
            foreach (XmlNode node in list)
            {
                if (node.Attributes != null && node.Attributes["id"] == null)
                    return "";
                if (node.Attributes == null ||
                    !node.Attributes["id"].Value.Equals(id, StringComparison.OrdinalIgnoreCase)) continue;
                var testMethod = GetNodeIndex(node, "TestMethod");
                var xmlAttributeCollection = node.ChildNodes[testMethod].Attributes;
                if (xmlAttributeCollection != null)
                    return xmlAttributeCollection[NodeName.ClassName].Value.Split(',')[0];
            }
            return string.Empty;
        }

        public List<string> GetOwners(XmlNodeList list, string id)
        {
            var ret_list = new List<string>();
            foreach (XmlNode node in list)
            {
                if (node.Attributes != null && node.Attributes["id"] == null)
                {
                    ret_list.Add("");
                    return ret_list;
                }
                if (node.Attributes == null ||
                    !node.Attributes["id"].Value.Equals(id, StringComparison.OrdinalIgnoreCase)) continue;

                foreach (XmlNode child1 in node.ChildNodes)
                {
                    if(child1.Name == "Owners")
                    {
                        foreach (XmlNode child2 in child1.ChildNodes)
                        {
                            if (child2.Name == "Owner")
                            {
                                var xmlAttributeCollection = child2.Attributes;
                                if ((xmlAttributeCollection != null) && (xmlAttributeCollection["name"] != null))
                                {
                                    string loc_val = xmlAttributeCollection["name"].Value;
                                    var loc_val_arr = loc_val.Split('|');
                                    foreach(string val_entry in loc_val_arr)
                                    {
                                        ret_list.Add(val_entry.Trim());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if(ret_list.Count == 0)
                ret_list.Add("");

            return ret_list;
        }
    }
}
