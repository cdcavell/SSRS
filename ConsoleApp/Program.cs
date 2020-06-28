using SSRS.NetCore.ClassLibrary;
using System;
using System.Collections.Generic;
using System.IO;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string reportExecution2005EndPointUrl = "https://laptop-v146j9kf/ReportServer/ReportExecution2005.asmx?wsdl";
            string reportService2005EndPointUrl = "https://laptop-v146j9kf/ReportServer/ReportService2005.asmx?wsdl";
            string reportPath = "/ReportProject/TestReport";
            byte[] result;

            Report report = new Report(reportExecution2005EndPointUrl, reportService2005EndPointUrl);

            // List Reports
            List<string[]> reportListing = report.Listing("/ReportProject");
            foreach (string[] item in reportListing)
            {
                Console.WriteLine(string.Format("[Name]: {0} [Description]: {1} [Path]: {2}",
                    item[0],
                    item[1],
                    item[2]
                ));
            }

            // Write "PDF"
            result = report.Render(reportPath).Result;
            WriteFile(result, report.Extension);

            // Write "EXCEL"
            result = report.Render(reportPath, "EXCEL").Result;
            WriteFile(result, report.Extension);

            // Write "WORD"
            result = report.Render(reportPath, "WORD").Result;
            WriteFile(result, report.Extension);

            // Write "XML"
            result = report.Render(reportPath, "XML").Result;
            WriteFile(result, report.Extension);

            // Write "CSV"
            result = report.Render(reportPath, "CSV").Result;
            WriteFile(result, report.Extension);

            // Write "IMAGE"
            result = report.Render(reportPath, "IMAGE").Result;
            WriteFile(result, report.Extension);

            // Write "HTML4.0"
            result = report.Render(reportPath, "HTML4.0").Result;
            WriteFile(result, report.Extension);

            // Write "MHTML"
            result = report.Render(reportPath, "MHTML").Result;
            WriteFile(result, report.Extension);

            Console.WriteLine("Done!");
        }

        private static void WriteFile(byte[] byteStream, string fileExtention)
        {
            string file = $"D:\\Temp\\TestFile." + fileExtention;

            using (var fs = File.OpenWrite(file))
            using (var sw = new StreamWriter(fs))
            {
                fs.Write(byteStream);
            }
        }
    }
}
