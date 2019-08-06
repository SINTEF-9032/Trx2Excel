﻿using System;
using System.Collections.Generic;

using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using Trx2Excel.Model;
using Trx2Excel.Setting;

namespace Trx2Excel.ExcelUtils
{
    public class ExcelWriter
    {
        private string FileName { get; set; }
        public ExcelWriter(string fileName)
        {
            FileName = fileName;
        }

        public void WriteToExcel(SortedDictionary<string, SortedDictionary<string, UnitTestResult>> total_results)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(FileName)))
            {
                var sheet = package.Workbook.Worksheets.Add("TotalResults");
                sheet = CreateHeader(sheet);
                var i = 2;
                var total_keys = total_results.Keys;
                foreach (var total_key in total_keys)
                {
                    var result_list = total_results[total_key];
                    var result_keys = result_list.Keys;
                    foreach (var result_key in result_keys)
                    {
                        var result = result_list[result_key];
                        sheet.Cells[i, 1].Value = result.Owner;
                        sheet.Cells[i, 1].AutoFitColumns();
                        sheet.Cells[i, 2].Value = result.NameSpace;
                        sheet.Cells[i, 2].AutoFitColumns();
                        sheet.Cells[i, 3].Value = result.TestName;
                        sheet.Cells[i, 3].AutoFitColumns();
                        sheet.Cells[i, 4].Value = result.Outcome;
                        sheet.Cells[i, 4].AutoFitColumns();
                        sheet.Cells[i, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[i, 4].Style.Fill.BackgroundColor.SetColor(
                            result.Outcome.Equals(TestOutcome.Failed.ToString(), StringComparison.OrdinalIgnoreCase) ?
                            Color.Red :
                            Color.ForestGreen);
                        sheet.Cells[i, 5].Value = result.Message;
                        sheet.Cells[i, 6].Value = result.StrackTrace;
                        sheet.Cells[i, 7].Value = result.AllOwnersString;
                        sheet.Cells[i, 8].Value = result.FileName;
                        i++;
                    }
                    i++;
                }
                package.Save();
            }
            
        }

        public void AddChart(int pass, int fail, int skip)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(FileName)))
            {
                var sheet = package.Workbook.Worksheets.Add("Result Chart");
                var chart = sheet.Drawings.AddChart("Result Chart", eChartType.Pie);
                var barChart = sheet.Drawings.AddChart("Result Bar Chart", eChartType.BarStacked);

                sheet.Cells["A2"].Value = "Passed";
                sheet.Cells["A2"].AutoFitColumns();
                sheet.Cells["A3"].Value = "Failed";
                sheet.Cells["A3"].AutoFitColumns();
                sheet.Cells["A4"].Value = "Skipped";
                sheet.Cells["A4"].AutoFitColumns();
                sheet.Cells["B2"].Value = pass;
                sheet.Cells["B3"].Value = fail;
                sheet.Cells["B4"].Value = skip;
                sheet.Cells["A2,A3,A4,B2,B3,B4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A2,B2"].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
                sheet.Cells["A3,B3"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                sheet.Cells["A4,B4"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                chart.Title.Text = "Test Result Pie Chart";
                chart.SetPosition(1, 0, 3, 0);
                var ser = chart.Series.Add("B2:B4", "A2:A4");
                ser.Header = "Count";

                barChart.Title.Text = "Test Result Bar Chart";
                barChart.SetPosition(14, 0, 3, 0);
                var barSer = barChart.Series.Add("B2:B4", "A2:A4");
                barSer.Header = "Count";
                package.Save();
            }
        }

        public ExcelWorksheet CreateHeader(ExcelWorksheet sheet)
        {
            string[] header = {"Owner", "Name Space", "Test Name", "Status", "Exception Message", "Stack Trace", "All owners", "File Name" };
            for (var i = 0; i < header.Length; i++)
            {
                sheet.Cells[1, i + 1].Value = header[i];
                sheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[1, i + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
                sheet.Cells[1, i + 1].AutoFitColumns();

            }
            return sheet;
        }
    }
}
