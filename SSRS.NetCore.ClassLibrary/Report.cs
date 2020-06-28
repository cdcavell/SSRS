using ReportExecution;
using ReportService;
using System;
using System.Collections.Generic;
using System.ServiceModel;
using System.Threading.Tasks;

namespace SSRS.NetCore.ClassLibrary
{
    public class Report
    {
        public string Extension;
        public string MimeType;
        public string Encoding;
        public string[] StreamIds;
        public List<KeyValuePair<string, string>> Warnings;

        private const string HistoryId = null;

        private List<ReportExecution.ParameterValue> ReportParameters;
        private ReportExecutionServiceSoapClient RSExecutionClient;
        private ReportingService2005SoapClient RSServiceClient;

        public Report(string reportExecution2005EndPointUrl, string reportService2005EndPointUrl)
        {
            ReportParameters = new List<ReportExecution.ParameterValue>();
            RSExecutionClient = CreateExecutionClient(reportExecution2005EndPointUrl);
            RSServiceClient = CreateServiceClient(reportService2005EndPointUrl);
        }

        public void AddParameters(List<KeyValuePair<string, string>> reportParameters)
        {
            foreach (KeyValuePair<string, string> kvp in reportParameters)
            {
                ReportParameters.Add(new ReportExecution.ParameterValue() { Name = kvp.Key, Value = kvp.Value });
            }
        }


        // https://docs.microsoft.com/en-us/dotnet/api/reportexecution2005.reportexecutionservice.render?view=sqlserver-2016
        // https://docs.microsoft.com/en-us/sql/reporting-services/customize-rendering-extension-parameters-in-rsreportserver-config?view=sql-server-ver15
        // https://medium.com/@yates.programmer/generating-an-ssrs-report-using-wcf-from-net-core-application-730e22886da3
        public async Task<byte[]> Render(string reportPath, string reportType = "PDF", string deviceInfo = null)
        {
            if (string.IsNullOrEmpty(reportPath))
                throw new Exception("Missing required reportPath parameter");

            reportType = reportType.ToUpper().Trim();
            switch (reportType)
            {
                case "PDF":
                case "EXCEL":
                case "WORD":
                case "XML":
                case "CSV":
                case "IMAGE":
                case "HTML4.0":
                case "MHTML":
                    break;
                default:
                    throw new Exception("Invalid reportType: " + reportType);
            }

            TrustedUserHeader trustedHeader = new TrustedUserHeader();
            LoadReportResponse loadReponse = await LoadReport(RSExecutionClient, trustedHeader, reportPath);

            if (ReportParameters.Count > 0)
            {
                await RSExecutionClient.SetExecutionParametersAsync(
                    loadReponse.ExecutionHeader,
                    trustedHeader,
                    ReportParameters.ToArray(),
                    "en-US"
                );
            }

            var renderRequest = new RenderRequest(loadReponse.ExecutionHeader, trustedHeader, reportType, deviceInfo);
            RenderResponse response = await RSExecutionClient.RenderAsync(renderRequest);

            if (response.Warnings != null)
            {
                foreach (ReportExecution.Warning warning in response.Warnings)
                {
                    Warnings.Add(
                        new KeyValuePair<string, string>(
                            warning.Code,
                            String.Format(
                                "Severity: {0} Object: {1} Message: {2}",
                                warning.Severity,
                                warning.ObjectName,
                                warning.Message
                            )
                        )
                    );
                }
            }

            Extension = response.Extension;
            MimeType = response.MimeType;
            Encoding = response.Encoding;
            StreamIds = response.StreamIds;

            return response.Result;
        }

        public List<string[]> Listing(string reportPath)
        {
            if (string.IsNullOrEmpty(reportPath))
                throw new Exception("Missing required reportPath parameter");

            List<string[]> listing = new List<string[]>();

            ReportService.ListChildrenResponse listChildrenResponse = null;
            listChildrenResponse = RSServiceClient.ListChildrenAsync(reportPath, false).Result;
            CatalogItem[] output = listChildrenResponse.CatalogItems;

            if (listChildrenResponse.CatalogItems.Length > 0)
            {
                foreach (CatalogItem item in listChildrenResponse.CatalogItems)
                {
                    string[] reportItem = new string[3];
                    reportItem[0] = item.Name ?? string.Empty;
                    reportItem[1] = item.Description ?? string.Empty;
                    reportItem[2] = item.Path ?? string.Empty;
                    
                    listing.Add(reportItem);
                }
            }

            return listing;
        }

        private ReportExecutionServiceSoapClient CreateExecutionClient(string reportExecution2005EndPointUrl)
        {
            BasicHttpsBinding rsBinding = new BasicHttpsBinding();
            rsBinding.Security.Mode = BasicHttpsSecurityMode.Transport;
            rsBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;

            // So we can download reports bigger than 64 KBytes
            // See https://stackoverflow.com/questions/884235/wcf-how-to-increase-message-size-quota
            rsBinding.MaxBufferPoolSize = 20000000;
            rsBinding.MaxBufferSize = 20000000;
            rsBinding.MaxReceivedMessageSize = 20000000;

            EndpointAddress rsEndpointAddress = new EndpointAddress(reportExecution2005EndPointUrl);
            ReportExecutionServiceSoapClient rsClient = new ReportExecutionServiceSoapClient(rsBinding, rsEndpointAddress);

            // Set user name and password
            rsClient.ClientCredentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;

            return rsClient;
        }

        private ReportingService2005SoapClient CreateServiceClient(string reportService2005EndPointUrl)
        {
            BasicHttpsBinding rsBinding = new BasicHttpsBinding();
            rsBinding.Security.Mode = BasicHttpsSecurityMode.Transport;
            rsBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;

            // So we can download reports bigger than 64 KBytes
            // See https://stackoverflow.com/questions/884235/wcf-how-to-increase-message-size-quota
            rsBinding.MaxBufferPoolSize = 20000000;
            rsBinding.MaxBufferSize = 20000000;
            rsBinding.MaxReceivedMessageSize = 20000000;

            EndpointAddress rsEndpointAddress = new EndpointAddress(reportService2005EndPointUrl);
            ReportingService2005SoapClient rsClient = new ReportingService2005SoapClient(rsBinding, rsEndpointAddress);

            // Set user name and password
            rsClient.ClientCredentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;

            return rsClient;
        }

        private async Task<LoadReportResponse> LoadReport(ReportExecutionServiceSoapClient rs, TrustedUserHeader trustedHeader, string reportPath)
        {
            // Get the report and set the execution header.
            // Failure to set the execution header will result in this error: "The session identifier is missing. A session identifier is required for this operation."
            // See https://social.msdn.microsoft.com/Forums/sqlserver/en-US/17199edb-5c63-4815-8f86-917f09809504/executionheadervalue-missing-from-reportexecutionservicesoapclient
            LoadReportResponse loadReponse = await rs.LoadReportAsync(trustedHeader, reportPath, HistoryId);

            return loadReponse;
        }
    }
}
