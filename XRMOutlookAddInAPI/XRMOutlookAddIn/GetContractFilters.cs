
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
    public static class GetContractFilters
    {
        public static string resourceId = "https://graph.microsoft.com";
        public static string tenantId = "70aa9dc9-726c-4d05-88f3-519ef4a1f1ac";
        public static string authString = "https://login.microsoftonline.com/" + tenantId;
        public static string upn = string.Empty;
        public static string clientId = "001bf6ce-45f9-4af4-bd57-ec96ea220e21";
        public static string clientSecret = "LnLN95vmqecwdaMv5AUq54g7uO3vMKjmvtJU5jlTAAo=";
        private static HttpClient _sharedHttpClient = new HttpClient();

        [FunctionName("GetContractFilters")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequestMessage req, ILogger log)
        {
            log.LogInformation("Function fetchuserdetails started");
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

                string body = "{\"requests\": [{\"url\": \"/sites/root/lists('Clients')/items?expand=fields(select=ClientName)&select=id,fields\",\"method\": \"GET\",\"id\": \"1\"},{\"url\": \"/sites/root/lists('Stakeholders')/items?expand=fields(select=StakeholderName)&select=id,fields\",\"method\": \"GET\",\"id\": \"2\"},{\"url\": \"/sites/root/lists('Status')/items?expand=fields(select=Title)&select=id,fields\",\"method\": \"GET\",\"id\": \"3\"}]}";
                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("POST"), requestUrl);
                requestMsg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                requestMsg.Content = new StringContent(body, Encoding.UTF8, "application/json");

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
                            Clients.Add(string.Format("{0},{1}",val.fields.ClientName,val.id));
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
                            Stakeholders.Add(string.Format("{0},{1}", val.fields.StakeholderName, val.id));
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