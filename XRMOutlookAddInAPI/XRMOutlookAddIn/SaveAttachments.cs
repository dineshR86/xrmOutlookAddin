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
using Newtonsoft.Json.Linq;
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
    public static class SaveAttachments
    {
        private static HttpClient _sharedHttpClient = new HttpClient();

        [FunctionName("SaveAttachments")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequestMessage req,ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            //Getting the Application settings
            string resourceId = Environment.GetEnvironmentVariable("ResourceId", EnvironmentVariableTarget.Process);
            string tenantid = Environment.GetEnvironmentVariable("TenantId", EnvironmentVariableTarget.Process);
            string authString = Environment.GetEnvironmentVariable("AuthString", EnvironmentVariableTarget.Process) + tenantid;
            string clientId = Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
            string clientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
            string ContractDriveID= Environment.GetEnvironmentVariable("ContractDriveID", EnvironmentVariableTarget.Process);
            string host = Environment.GetEnvironmentVariable("Host", EnvironmentVariableTarget.Process);

            try
            {
                AttachmentProps props = await req.Content.ReadAsAsync<AttachmentProps>();
                string requesturl = string.Format("https://graph.microsoft.com/v1.0/users/{0}/messages/{1}/attachments", props.UserId, props.MessageId);
                var authenticationContext = new AuthenticationContext(authString, false);
                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(resourceId, clientCred);
                string token = authenticationResult.AccessToken;
                //HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("GET"), requesturl);
                //requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                
                //HttpResponseMessage response = _sharedHttpClient.SendAsync(requestMsg).Result;

                //var content = await response.Content.ReadAsStringAsync();
                await CheckForFolder(string.Format("{0}-{1}", props.ItemTitle, props.ItemID), ContractDriveID, token);
                dynamic items = JsonConvert.DeserializeObject<RootAttachment>("");

                foreach(var item in items.value)
                {
                    await UploadFileToLibrary(item.contentBytes, item.Name, token, string.Format("{0}-{1}", props.ItemTitle, props.ItemID), ContractDriveID);
                }

                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("Success") };
            }
            catch(Exception ex)
            {
                log.LogError(string.Format("Exception! '{0}'.", ex));
                return req.CreateResponse(HttpStatusCode.InternalServerError, new { summary = "Error" });
            }
        }

        private static async Task<Boolean> UploadFileToLibrary(string data,string docName,string accessToken,string folderName,string driveid)
        {
            string uploadUri = string.Format("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content", driveid, folderName,docName);
            HttpRequestMessage uploadRequest = new HttpRequestMessage(new HttpMethod("PUT"), uploadUri);
            uploadRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            Byte[] byteArray = Convert.FromBase64String(data);
            uploadRequest.Content = new ByteArrayContent(byteArray);
            uploadRequest.Content.Headers.Add("Content-Length", byteArray.Length.ToString());
            using (var uploadresponse = await _sharedHttpClient.SendAsync(uploadRequest))
            {
                if (!uploadresponse.IsSuccessStatusCode)
                {
                    throw new Exception(uploadresponse.ReasonPhrase);
                }
            }

            return true;
        }

        private static async Task<Boolean> CheckForFolder(string folderName,string driveid,string accessToken)
        {
            string folderUri = string.Format("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}", driveid, folderName);
            HttpRequestMessage folderRequest = new HttpRequestMessage(new HttpMethod("GET"), folderUri);
            folderRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            
            using (var folderresponse = await _sharedHttpClient.SendAsync(folderRequest))
            {
                if (folderresponse.IsSuccessStatusCode)
                {
                    var json = JObject.Parse(await folderresponse.Content.ReadAsStringAsync());
                    json.GetValue("id");
                    return true;
                }
                else
                {
                    string createfolderUri= string.Format("https://graph.microsoft.com/v1.0/drives/{0}/root/children", driveid);
                    dynamic cfolder = new JObject();
                    cfolder.name = folderName;
                    cfolder.folder = new JObject();
                    //cfolder.@microsoft.graph.conflictBehavior = "rename";
                    HttpRequestMessage cfolderrequest = new HttpRequestMessage(new HttpMethod("POST"), createfolderUri);
                    cfolderrequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    cfolderrequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    cfolderrequest.Content = new StringContent(cfolder.ToString(), Encoding.UTF8, "application/json");
                    var cFolderresponse = await _sharedHttpClient.SendAsync(cfolderrequest);
                    return true;
                }
            }
        }
    }

    internal class AttachmentProps
    {
        public string MessageId { get; set; }

        public string UserId { get; set; }

        public string ItemTitle { get; set; }

        public string ItemID { get; set; }
    }

    internal class Attachment
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string ContentType { get; set; }
        public int size { get; set; }
        public string contentBytes { get; set; }
    }

    internal class RootAttachment
    {
        public List<Attachment> value;
    }

}
