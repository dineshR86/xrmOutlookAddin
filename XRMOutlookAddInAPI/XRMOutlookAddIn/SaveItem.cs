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
    public static class SaveItem
    {
        public static string resourceId = "https://graph.microsoft.com";
        public static string tenantId = "70aa9dc9-726c-4d05-88f3-519ef4a1f1ac";
        public static string authString = "https://login.microsoftonline.com/" + tenantId;
        public static string upn = string.Empty;
        public static string clientId = "001bf6ce-45f9-4af4-bd57-ec96ea220e21";
        public static string clientSecret = "LnLN95vmqecwdaMv5AUq54g7uO3vMKjmvtJU5jlTAAo=";
        private static HttpClient _sharedHttpClient = new HttpClient();
        private static string host = "oaktondidata.sharepoint.com";

        [FunctionName("SaveItem")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequestMessage req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            try
            {
                GetData fields = await req.Content.ReadAsAsync<GetData>();
                MailData postFields = new MailData()
                {
                    Subject = fields.Subject,
                    To = fields.To,
                    Message = fields.Message,
                    From = fields.From,
                    conversationId = fields.ConversationId,
                    conversationTopic = fields.ConversationTopic,
                    received = fields.Received,
                    itemid = fields.itemid,
                    listid = fields.listid
                };
                PostData data = new PostData();
                data.fields = postFields;
                var posdata = JsonConvert.SerializeObject(data);

                string rel = new Uri(fields.sitecollectionUrl).AbsolutePath;
                string siteurl = "";
                siteurl = rel == "/" ? host : string.Format("{0}:{1}", host, rel);
                string listname = fields.listname;
                string requesturl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}:/lists('{1}')/items", siteurl, listname);

                var authenticationContext = new AuthenticationContext(authString, false);
                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;
                if (!string.IsNullOrEmpty(token))
                {
                    log.LogInformation("successfully obtained access token");
                }

                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("POST"), requesturl);
                requestMsg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                requestMsg.Content = new StringContent(posdata, Encoding.UTF8, "application/json");
                HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;
                var content = await response.Content.ReadAsStringAsync();

                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("Success") };
            }
            catch (Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return req.CreateResponse(HttpStatusCode.InternalServerError, new { summary = "Error" });
            }
        }
    }

    internal class MailData
    {
        [JsonProperty(PropertyName = "Subject")]
        public string Subject { get; set; }
        [JsonProperty(PropertyName = "To")]
        public string To { get; set; }
        [JsonProperty(PropertyName = "Message")]
        public string Message { get; set; }
        [JsonProperty(PropertyName = "From")]
        public string From { get; set; }

        [JsonProperty(PropertyName ="Title")]
        public string conversationId { get; set; }

        [JsonProperty(PropertyName = "Conversation_x0020_Topic")]
        public string conversationTopic { get; set; }
        [JsonProperty(PropertyName = "Received")]
        public string received { get; set; }
        [JsonProperty(PropertyName = "Related_x0020_Item_x0020_Id")]
        public string itemid { get; set; }
        [JsonProperty(PropertyName = "Related_x0020_Item_x0020_List_x0")]
        public string listid { get; set; }

    }

    internal class GetData
    {
        public string Subject { get; set; }
        public string To { get; set; }
        public string Message { get; set; }
        public string From { get; set; }
        public string ConversationId { get; set; }
        public string ConversationTopic { get; set; }
        public string Received { get; set; }
        public string itemid { get; set; }
        public string listid { get; set; }
        public string sitecollectionUrl { get; set; }
        public string listname { get; set; }
    }

    internal class PostData
    {
        public MailData fields { get; set; }
    }
}
