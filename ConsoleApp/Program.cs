using SSRS.NetCore.ClassLibrary;
using System;
using System.IO;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string reportExecution2005EndPointUrl = "https://laptop-v146j9kf/ReportServer/ReportExecution2005.asmx?wsdl";
            string reportPath = "/ReportProject/TestReport";
            byte[] result;

            Report report = new Report(reportExecution2005EndPointUrl, reportPath);

            // Write "PDF"
            result = report.Render().Result;
            WriteFile(result, report.Extension);

            // Write "EXCEL"
            result = report.Render("EXCEL").Result;
            WriteFile(result, report.Extension);

            // Write "WORD"
            result = report.Render("WORD").Result;
            WriteFile(result, report.Extension);

            // Write "XML"
            result = report.Render("XML").Result;
            WriteFile(result, report.Extension);

            // Write "CSV"
            result = report.Render("CSV").Result;
            WriteFile(result, report.Extension);

            // Write "IMAGE"
            result = report.Render("IMAGE").Result;
            WriteFile(result, report.Extension);

            // Write "HTML4.0"
            result = report.Render("HTML4.0").Result;
            WriteFile(result, report.Extension);

            // Write "MHTML"
            result = report.Render("MHTML").Result;
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
