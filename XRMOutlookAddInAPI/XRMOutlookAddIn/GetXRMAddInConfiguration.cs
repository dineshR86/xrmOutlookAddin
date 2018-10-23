
using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace XRMOutlookAddIn
{
    public static class GetXRMAddInConfiguration
    {
        public static string resourceId = "https://graph.microsoft.com";
        public static string tenantId = "70aa9dc9-726c-4d05-88f3-519ef4a1f1ac";
        public static string authString = "https://login.microsoftonline.com/" + tenantId;
        public static string upn = string.Empty;
        public static string clientId = "001bf6ce-45f9-4af4-bd57-ec96ea220e21";
        public static string clientSecret = "LnLN95vmqecwdaMv5AUq54g7uO3vMKjmvtJU5jlTAAo=";
        private static HttpClient _sharedHttpClient = new HttpClient();

        [FunctionName("GetXRMAddInConfiguration")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequestMessage req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
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

                string requestUrl = "https://graph.microsoft.com/v1.0/sites/root/lists('XRMAddinConfiguration')/items?expand=fields(select=ConfigKey,ConfigValue)";
                log.LogInformation(string.Format("About to hit Graph endpoint: '{0}'.", requestUrl));

                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("GET"), requestUrl);
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                var content =  await response.Content.ReadAsStringAsync();
                dynamic items = JsonConvert.DeserializeObject<RootValue>(content);
                ConfigData config = new ConfigData();
                foreach (var item in items.value)
                {
                    switch (item.fields.ConfigKey)
                    {
                        case "SiteCollections":
                            config.SiteCollectionUrls = item.fields.ConfigValue;
                            break;
                        case "Lists":
                            config.Lists = item.fields.ConfigValue;
                            break;
                        case "CaseStatusFilter":
                            config.CaseStatusFilter = item.fields.ConfigValue;
                            break;
                        case "ProjectStatusFilter":
                            config.ProjectStatusFilter = item.fields.ConfigValue;
                            break;
                        default:
                            break;
                    }
                }
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonConvert.SerializeObject(config,Formatting.Indented),Encoding.UTF8,"application/json")};
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return req.CreateResponse(HttpStatusCode.InternalServerError, new { summary = "Error" });
            }

        }

        
    }

    internal class RootValue
    {
        public List<Value> value;
    }

    internal class Value
    {
        public string id { get; set; }
        public string webUrl { get; set; }
        public Fields fields { get; set; }
    }

    internal class Fields
    {
        public string ConfigKey { get; set; }
        public string ConfigValue { get; set; }
        public string ClientName { get; set; }
        public string StakeholderName { get; set; }
        public string Title { get; set; }
    }

    internal class ConfigData
    {
        public string SiteCollectionUrls { get; set; }
        public string Lists { get; set; }
        public string CaseStatusFilter { get; set; }
        public string ProjectStatusFilter { get; set; }

    }
}