
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace XRMOutlookAddIn
{
    public static class GetListItems
    {
        public static string resourceId = "https://graph.microsoft.com";
        public static string tenantId = "70aa9dc9-726c-4d05-88f3-519ef4a1f1ac";
        public static string authString = "https://login.microsoftonline.com/" + tenantId;
        public static string upn = string.Empty;
        public static string clientId = "001bf6ce-45f9-4af4-bd57-ec96ea220e21";
        public static string clientSecret = "LnLN95vmqecwdaMv5AUq54g7uO3vMKjmvtJU5jlTAAo=";
        private static HttpClient _sharedHttpClient = new HttpClient();
        private static string host = "oaktondidata.sharepoint.com";


        [FunctionName("GetListItems")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function for fetching list items");

            //string sitecollection = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0).Value;
            string sitecollection = req.Query["sc"];
            string listname = req.Query["list"];
            string fieldname = req.Query["ff"];
            string fieldvalue = req.Query["val"];


            try
            {
                string rel = new Uri(sitecollection).AbsolutePath;
                string requestUrl = "";
                // sample graph api call with filter https://graph.microsoft.com/v1.0/sites/oaktondidata.sharepoint.com/lists('XRMCases')/items?expand=fields(select=Title,StatusLookupId)&filter=fields/StatusLookupId eq '3'
                if (rel == "/")
                {
                    requestUrl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}/lists('{1}')/items", host, listname);
                }
                else
                {
                     requestUrl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}:{1}/lists('{2}')/items", host, rel,listname);
                }
                
                var authenticationContext = new AuthenticationContext(authString, false);

                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;

                if (!string.IsNullOrEmpty(token))
                {
                    log.LogInformation("successfully obtained access token");
                }


                log.LogInformation(string.Format("About to hit Graph endpoint: '{0}'.", requestUrl));

                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("GET"), requestUrl);
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                var content = await response.Content.ReadAsStringAsync();
                
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(content) };
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent("Error") };
            }
        }
    }
}
