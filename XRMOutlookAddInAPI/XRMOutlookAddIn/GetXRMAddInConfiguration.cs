
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
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
using Microsoft.AspNetCore.Http;

namespace XRMOutlookAddIn
{
    public static class GetXRMAddInConfiguration
    {   
        private static HttpClient _sharedHttpClient = new HttpClient();
        public static string ClientId = string.Empty;
        public static string ClientSecret = string.Empty;
        public static string TenantId = string.Empty;
        public static string Host = string.Empty;

        [FunctionName("GetXRMAddInConfiguration")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            try
            {   
                //hardcoding the domain which will be removed. Need to get the domain as part of the email.
                var domain = req.Query["domain"];
                Host = $"{domain}.sharepoint.com";
                var keyVaultName = Environment.GetEnvironmentVariable("KeyVaultName", EnvironmentVariableTarget.Process);
                var azureServiceTokenProvider = new AzureServiceTokenProvider();
                var keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));
                string connection= (await keyVaultClient.GetSecretAsync($"https://{keyVaultName}.vault.azure.net/secrets/{domain + "Connection"}")).Value;
                ClientId = connection.Split(';')[1];
                ClientSecret = connection.Split(';')[2];
                TenantId = connection.Split(';')[0];

                //Getting the Application settings
                string resourceId = Environment.GetEnvironmentVariable("ResourceId", EnvironmentVariableTarget.Process);
                string authString = Environment.GetEnvironmentVariable("AuthString", EnvironmentVariableTarget.Process) + TenantId;

                var authenticationContext = new AuthenticationContext(authString, false);
                ClientCredential clientCred = new ClientCredential(ClientId, ClientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;

                if (!string.IsNullOrEmpty(token))
                {
                    log.LogInformation("successfully obtained access token");
                }

                string requestUrl = "https://graph.microsoft.com/v1.0/sites/root/lists('XRMAddinConfiguration')/items?expand=fields(select=ConfigKey,ConfigValue)";

                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("GET"), requestUrl);
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                if (response.IsSuccessStatusCode) { 
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
                else
                {
                    throw new Exception("Error while fetching the XRMConfiguration. Please contact the administrator.");
                }
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                //return req.CreateResponse(HttpStatusCode.InternalServerError, new { summary = ex.Message });
                return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent(ex.Message) };
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
        public string Client_x0020_Name { get; set; }
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
