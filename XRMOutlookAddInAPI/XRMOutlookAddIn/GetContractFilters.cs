
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
using Newtonsoft.Json.Linq;


namespace XRMOutlookAddIn
{
    public static class GetContractFilters
    {
        private static HttpClient _sharedHttpClient = new HttpClient();

        [FunctionName("GetContractFilters")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequestMessage req, ILogger log)
        {
            log.LogInformation("Function GetContractFilters started");
            //Getting the Application settings
            string resourceId = Environment.GetEnvironmentVariable("ResourceId", EnvironmentVariableTarget.Process);
            string tenantid = Environment.GetEnvironmentVariable("TenantId", EnvironmentVariableTarget.Process);
            string authString = Environment.GetEnvironmentVariable("AuthString", EnvironmentVariableTarget.Process) + tenantid;
            string clientId = Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
            string clientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
            string host = Environment.GetEnvironmentVariable("Host", EnvironmentVariableTarget.Process);

            //hardcoding of the url needs to be removed
            string rel = new Uri("https://cloudmission.sharepoint.com/sites/xrm").AbsolutePath;
            string siteurl = "";
            siteurl = rel == "/" ? host : string.Format("{0}:{1}:", host, rel);

            try
            {
                var authenticationContext = new AuthenticationContext(authString, false);

                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;

                if (!string.IsNullOrEmpty(token))
                {
                    log.LogInformation("successfully obtained access token");
                }

                string requestUrl = "https://graph.microsoft.com/v1.0/$batch";
                log.LogInformation(string.Format("About to hit Graph endpoint: '{0}'.", requestUrl));

                JObject req1 = new JObject{
                {"id","1"},
                {"method","GET"},
                {"url",string.Format("/sites/{0}/lists('Clients')/items?expand=fields(select=Title)&select=id,fields",siteurl)}
            };
                JObject req2 = new JObject{
                {"id","2"},
                {"method","GET"},
                {"url",string.Format("/sites/{0}/lists('Stakeholders')/items?expand=fields(select=Title)&select=id,fields",siteurl)}
            };
                JObject req3 = new JObject{
                {"id","3"},
                {"method","GET"},
                {"url",string.Format("/sites/{0}/lists('Case Statuses')/items?expand=fields(select=Title)&select=id,fields",siteurl)}
            };
                JArray a = new JArray();
                a.Add(req1); a.Add(req2); a.Add(req3);

                JObject o = new JObject();
                o["requests"] = a;
                
                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("POST"), requestUrl);
                requestMsg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                requestMsg.Content = new StringContent(o.ToString(), Encoding.UTF8, "application/json");

                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                var content = await response.Content.ReadAsStringAsync();
                dynamic items = JsonConvert.DeserializeObject<RootObject>(content);
                List<string> Clients = new List<string>();
                List<string> Stakeholders = new List<string>();
                List<string> Status = new List<string>();
                foreach (var item in items.responses)
                {
                    if (item.id == "1")
                    {
                        foreach (var val in item.body.value)
                        {
                            Clients.Add(string.Format("{0},{1}",val.fields.Title, val.id));
                        }
                    }else if (item.id == "3")
                    {
                        foreach (var val in item.body.value)
                        {
                            Status.Add(string.Format("{0},{1}", val.fields.Title, val.id));
                        }
                    }
                    else
                    {
                        foreach (var val in item.body.value)
                        {
                            Stakeholders.Add(string.Format("{0},{1}", val.fields.Title, val.id));
                        }
                    }
                }

                FilterObject filterdata = new FilterObject();
                filterdata.Clients = Clients;
                filterdata.Stakeholders = Stakeholders;
                filterdata.Status = Status;
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonConvert.SerializeObject(filterdata, Formatting.Indented), Encoding.UTF8, "application/json") };
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return req.CreateResponse(HttpStatusCode.InternalServerError, new { summary = "Error" });
            }
        }
    }

    internal class Body
    {
        public List<Value> value { get; set; }
    }

    internal class Respons
    {
        public string id { get; set; }
        public int status { get; set; }
        public Body body { get; set; }
    }

    internal class RootObject
    {
        public List<Respons> responses { get; set; }
    }

    internal class FilterObject
    {
        public List<string> Clients { get; set; }
        public List<string> Stakeholders { get; set; }
        public List<string> Status { get; set; }
    }
}
