// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class GroupsController
    {
        public async Task<IEnumerable<AADIdentity>> GetGroups(HttpClient graphHttpClient)
        {
            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            HttpResponseMessage response = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/groups");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"List groups graph call failed: {response.StatusCode}-{responseMsg}");
            }

            GraphDataSet<AADIdentity> dataSet = JsonConvert.DeserializeObject<GraphDataSet<AADIdentity>>(responseMsg);
            return dataSet.Value;
        }

        public async Task<AADIdentity> CreateGroup(HttpClient graphHttpClient, string groupName)
        {
            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            if (string.IsNullOrWhiteSpace(groupName))
            {
                throw new ArgumentException(nameof(groupName));
            }

            string groupDetail =
                "{\"description\": \"group for #groupName#\",\"displayName\": \"#groupName#\",\"groupTypes\": [\"Unified\"],\"mailEnabled\": true,\"mailNickname\": \"#mail#\",\"securityEnabled\": false, \"owners@odata.bind\": [\"https://graph.microsoft.com/v1.0/users/#owner#\"]}";
            groupDetail = groupDetail.Replace("#groupName#", groupName)
                .Replace("#mail#", groupName.Replace(" ", string.Empty).Replace(",", string.Empty))
                .Replace("#owner#", Settings.NewOwnerId);

            HttpContent content = new StringContent(groupDetail, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PostAsync($"{Settings.GraphBaseUri}/groups", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new AutoTeamsStructureException($"Create group graph call failed for {groupName}: {response.StatusCode}-{responseMsg}");
            }

            return JsonConvert.DeserializeObject<AADIdentity>(responseMsg);
        }
    }
}
