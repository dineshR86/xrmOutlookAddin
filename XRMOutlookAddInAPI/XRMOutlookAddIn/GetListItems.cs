
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

        private static HttpClient _sharedHttpClient = new HttpClient();

        [FunctionName("GetListItems")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function for fetching list items");

            //Getting the Application settings
            string resourceId = Environment.GetEnvironmentVariable("ResourceId", EnvironmentVariableTarget.Process);
            string tenantid = Environment.GetEnvironmentVariable("TenantId", EnvironmentVariableTarget.Process);
            string authString = Environment.GetEnvironmentVariable("AuthString", EnvironmentVariableTarget.Process) + tenantid;
            string clientId = Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
            string clientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
            string host = Environment.GetEnvironmentVariable("Host", EnvironmentVariableTarget.Process);

            //string sitecollection = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0).Value;
            string sitecollection = req.Query["sc"];
            string listname = req.Query["list"];
            string fieldname = req.Query["ff"];
            string fieldvalue = req.Query["val"];
            string fieldname1 = req.Query["ff1"];
            string fieldvalue1 = req.Query["val1"];


            try
            {
                string rel = new Uri(sitecollection).AbsolutePath;
                string siteurl = "";
                // sample graph api call with filter https://graph.microsoft.com/v1.0/sites/oaktondidata.sharepoint.com/lists('XRMCases')/items?expand=fields(select=Title,StatusLookupId)&filter=fields/StatusLookupId eq '3'
                siteurl = rel == "/" ? host : string.Format("{0}:{1}:", host, rel);
                string requestUrl = "";
                if (!string.IsNullOrEmpty(fieldname1))
                {
                    requestUrl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?expand=fields(select=Title,{2})&filter=fields/{2} eq '{3}'and fields/{4} eq '{5}'&select=id,fields", siteurl, listname, fieldname, fieldvalue, fieldname1, fieldvalue1);
                }
                else
                {
                    requestUrl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?expand=fields(select=Title,{2})&filter=fields/{2} eq '{3}'&select=id,fields", siteurl, listname, fieldname, fieldvalue);
                }

                var authenticationContext = new AuthenticationContext(authString, false);

                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;

                if (!string.IsNullOrEmpty(token))
                {
                    log.LogInformation("successfully obtained access token");
                }

                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("GET"), requestUrl);
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    dynamic items = JsonConvert.DeserializeObject<RootValue>(content);
                    List<ItemData> datas = new List<ItemData>();
                    foreach (var item in items.value)
                    {
                        ItemData data = new ItemData();
                        data.ID = item.id;
                        data.Title = item.fields.Title;
                        datas.Add(data);
                    }

                    return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonConvert.SerializeObject(datas, Formatting.Indented), Encoding.UTF8, "application/json") };
                }
                else
                {
                    throw new Exception("Error while fetching the list items. Please contact the administrator.");
                }
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent(ex.Message) };
            }
        }
    }

    internal class ItemData
    {
        public string ID { get; set; }
        public string Title { get; set; }

    }
}
