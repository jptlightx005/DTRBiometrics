using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DTRFuncs
{
    public class DocumentExporter
    {
        public static void exportToExcel(Dictionary<String, Object> record)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            Dictionary<String, DateTime> timePeriod = record["Period"] as Dictionary<String, DateTime>;
            if (timePeriod != null)
            {
                var from = timePeriod["From"].ToString("MMMM d");
                var to = timePeriod["To"].ToString("d, yyyy");

                var periodText = from + " - " + to;

                workSheet.Cells[2, "B"] = periodText.ToUpper();
                workSheet.Cells[2, "H"] = periodText.ToUpper();

                workSheet.Range["B:B", System.Type.Missing].EntireColumn.ColumnWidth = 16;
                workSheet.Range["H:H", System.Type.Missing].EntireColumn.ColumnWidth = 16;
                var bf = workSheet.Range[workSheet.Cells[6, "B"], workSheet.Cells[5 + timePeriod["To"].Day, "F"]];
                var hl = workSheet.Range[workSheet.Cells[6, "H"], workSheet.Cells[5 + timePeriod["To"].Day, "L"]];
                bf.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                hl.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                for (int i = timePeriod["From"].Day; i <= timePeriod["To"].Day; i++)
                {
                    var cellDay = new DateTime(timePeriod["From"].Year, timePeriod["From"].Month, i);
                    var dayOfWeek = cellDay.ToString("dddd");
                    var dayCutted = dayOfWeek.Substring(0, 2);
                    workSheet.Cells[5 + i, "B"] = cellDay.ToString("MM/dd/yyyy") + " " + dayCutted;
                    workSheet.Cells[5 + i, "H"] = cellDay.ToString("MM/dd/yyyy") + " " + dayCutted;
                }

                var rowNumber = 8 + timePeriod["To"].Day; //38 for november
                workSheet.Range[workSheet.Cells[rowNumber, "B"], workSheet.Cells[rowNumber, "F"]].Merge();
                workSheet.Cells[rowNumber, "B"] = "I certify on my honor that the above is a true and correct report of the hours of work performed, records of which was made daily at the time of arrival and departure from Office.";

                workSheet.Range[workSheet.Cells[rowNumber, "H"], workSheet.Cells[rowNumber, "L"]].Merge();
                workSheet.Cells[rowNumber, "H"] = "I certify on my honor that the above is a true and correct report of the hours of work performed, records of which was made daily at the time of arrival and departure from Office.";

                var row3s = workSheet.Range[String.Format("{0}:{1}", rowNumber, rowNumber), System.Type.Missing];
                row3s.EntireRow.RowHeight = 45;
                row3s.WrapText = true;

                var signatureRowNumber = rowNumber + 3; //41
                var signLabelRowNumber = rowNumber + 4; //42
                var designationRowNumber = rowNumber + 5; //43
                var desgnLabelRowNumber = rowNumber + 6; //44

                for(int i = 3; i <= 13; i++)
                {
                    workSheet.Range[workSheet.Cells[rowNumber + i, "B"], workSheet.Cells[rowNumber + i, "F"]].Merge();
                    workSheet.Range[workSheet.Cells[rowNumber + i, "H"], workSheet.Cells[rowNumber + i, "L"]].Merge();
                }
                workSheet.Range[workSheet.Cells[signatureRowNumber, "B"], workSheet.Cells[signatureRowNumber, "F"]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Range[workSheet.Cells[signatureRowNumber, "H"], workSheet.Cells[signatureRowNumber, "L"]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                workSheet.Cells[signLabelRowNumber, "B"] = "Signature";
                workSheet.Cells[signLabelRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[signLabelRowNumber, "H"] = "Signature";
                workSheet.Cells[signLabelRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[designationRowNumber, "B"] = "Senior Administrative Assistant I";
                workSheet.Cells[designationRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[designationRowNumber, "H"] = "Senior Administrative Assistant I";
                workSheet.Cells[designationRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[desgnLabelRowNumber, "B"] = "Designation";
                workSheet.Cells[desgnLabelRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[desgnLabelRowNumber, "H"] = "Designation";
                workSheet.Cells[desgnLabelRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                var verifiedLabelRowNumber = rowNumber + 9;
                var accountantRowNumber = rowNumber + 12;
                var accountantLabelRowNumber = rowNumber + 13;

                workSheet.Cells[verifiedLabelRowNumber, "B"] = "Verified as to prescribed Office hours.";
                workSheet.Cells[verifiedLabelRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[verifiedLabelRowNumber, "H"] = "Verified as to prescribed Office hours.";
                workSheet.Cells[verifiedLabelRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                workSheet.Cells[accountantRowNumber, "B"] = "Joyfa C. Ramos".ToUpper();
                workSheet.Cells[accountantRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[accountantRowNumber, "B"].Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                workSheet.Cells[accountantRowNumber, "H"] = "Joyfa C. Ramos".ToUpper();
                workSheet.Cells[accountantRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[accountantRowNumber, "H"].Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;

                workSheet.Cells[accountantLabelRowNumber, "B"] = "Municipal Accountant";
                workSheet.Cells[accountantLabelRowNumber, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[accountantLabelRowNumber, "H"] = "Municipal Accountant";
                workSheet.Cells[accountantLabelRowNumber, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            String employee = record["Employee"] as String;
            if(employee != null)
            {
                workSheet.Cells[4, "B"] = employee.ToUpper();
                workSheet.Cells[4, "B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[4, "B"].Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                workSheet.Range[workSheet.Cells[4, "B"], workSheet.Cells[4, "F"]].Merge();

                workSheet.Cells[4, "H"] = employee.ToUpper();
                workSheet.Cells[4, "H"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[4, "H"].Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                workSheet.Range[workSheet.Cells[4, "H"], workSheet.Cells[4, "L"]].Merge();
            }

            List<Dictionary<String, Object>> timeLogs = record["TimeLogs"] as List<Dictionary<String, Object>>;
            if (timeLogs != null)
            {
                foreach(var timeLog in timeLogs)
                {
                    var current = (DateTime)timeLog["DateOfTheDay"];
                    int i = current.Day;
                    workSheet.Cells[5 + i, "C"] = timeLog["TimeInAM"];
                    workSheet.Cells[5 + i, "D"] = timeLog["TimeOutAM"];
                    workSheet.Cells[5 + i, "E"] = timeLog["TimeInPM"];
                    workSheet.Cells[5 + i, "F"] = timeLog["TimeOutPM"];

                    workSheet.Cells[5 + i, "I"] = timeLog["TimeInAM"];
                    workSheet.Cells[5 + i, "J"] = timeLog["TimeOutAM"];
                    workSheet.Cells[5 + i, "K"] = timeLog["TimeInPM"];
                    workSheet.Cells[5 + i, "L"] = timeLog["TimeOutPM"];
                }
            }

        }

        static string returnBlankIfNull(object data)
        {
            var filtered = data as String;

            return filtered != null ? filtered : "";
        }
    }
}
