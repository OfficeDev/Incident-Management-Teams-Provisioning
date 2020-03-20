// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class TeamsController
    {
        public async Task CreatedTeamsFromGroupidAsync(string groupId, HttpClient graphHttpClient)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            string teamDetail =
                "{  \"memberSettings\": {\"allowCreateUpdateChannels\": true},\"messagingSettings\": {\"allowUserEditMessages\": true,\"allowUserDeleteMessages\": true},\"funSettings\": {\"allowGiphy\": true,\"giphyContentRating\": \"strict\"}}";
            HttpContent content = new StringContent(teamDetail, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PutAsync($"{Settings.GraphBaseUri}/groups/{groupId}/team", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new AutoTeamsStructureException($"Create teams team graph call failed for {groupId}: {response.StatusCode}-{responseMsg}");
            }
        }

        public async Task<string> CreateChannelAsync(string groupId, string channelName, HttpClient graphHttpClient)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (string.IsNullOrWhiteSpace(channelName))
            {
                throw new ArgumentException(nameof(channelName));
            }

            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            HttpContent content = new StringContent(JsonConvert.SerializeObject(new { displayName = channelName, description = "Automatically generated channel" }), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PostAsync($"{Settings.GraphBaseUri}/teams/{groupId}/channels", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new AutoTeamsStructureException($"Create teams channel graph call failed for {groupId}->{channelName}: {response.StatusCode}-{responseMsg}");
            }

            return JsonConvert.DeserializeObject<AADIdentity>(responseMsg).Id;
        }

        public async Task UploadFileAsync(string groupId, string channelName, byte[] fileContent, string fileName, HttpClient graphHttpClient)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (string.IsNullOrWhiteSpace(channelName))
            {
                throw new ArgumentException(nameof(channelName));
            }

            if (fileContent == null)
            {
                throw new ArgumentException(nameof(fileContent));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException(nameof(fileName));
            }

            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            HttpContent content = new StreamContent(new MemoryStream(fileContent));
            HttpResponseMessage response = await graphHttpClient.PutAsync($"{Settings.GraphBaseUri}/groups/{groupId}/drive/root:/{channelName}/{fileName}:/content", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created && response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"Update file to channel graph call failed for {groupId}->{channelName}: {response.StatusCode}-{responseMsg}");
            }
        }

        public async Task AddCustomTabAsync(string groupId, string channelId, TeamsApp appInfo, HttpClient graphHttpClient)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (string.IsNullOrWhiteSpace(channelId))
            {
                throw new ArgumentException(nameof(channelId));
            }

            if (appInfo == null)
            {
                throw new ArgumentException(nameof(appInfo));
            }

            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            HttpContent content = new StringContent(JsonConvert.SerializeObject(appInfo), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PostAsync($"{Settings.GraphBaseUri}/teams/{groupId}/channels/{channelId}/tabs", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new AutoTeamsStructureException($"Create teams channel tab graph call failed for {groupId}->{channelId}: {response.StatusCode}-{responseMsg}");
            }
        }

        public async Task<IEnumerable<AADIdentity>> GetChannels(string groupId, HttpClient graphHttpClient)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            HttpResponseMessage response = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/teams/{groupId}/channels");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"Get teams channel failed for {groupId}: {responseMsg}");
            }

            GraphDataSet<AADIdentity> dataSet = JsonConvert.DeserializeObject<GraphDataSet<AADIdentity>>(responseMsg);
            return dataSet.Value;
        }
    }
}
