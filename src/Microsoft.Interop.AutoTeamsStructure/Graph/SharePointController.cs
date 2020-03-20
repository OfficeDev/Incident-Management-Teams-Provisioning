// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class SharePointController
    {
        public IEnumerable<JObject> GetListItems(HttpClient graphHttpClient, string groupId, string listName)
        {
            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            if (groupId == null)
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (listName == null)
            {
                throw new ArgumentException(nameof(listName));
            }

            string listId = GetListId(graphHttpClient, groupId, listName).GetAwaiter().GetResult();
            if (listId != null)
            {
                return GetListItemsById(graphHttpClient, groupId, listId).GetAwaiter().GetResult();
            }

            return null;
        }

        public string GetDocItemDownloadUrl(HttpClient graphHttpClient, string groupId, string listName,
            string itemId)
        {
            if (graphHttpClient == null)
            {
                throw new ArgumentException(nameof(graphHttpClient));
            }

            if (groupId == null)
            {
                throw new ArgumentException(nameof(groupId));
            }

            if (listName == null)
            {
                throw new ArgumentException(nameof(listName));
            }

            if (itemId == null)
            {
                throw new ArgumentException(nameof(itemId));
            }

            string listId = GetListId(graphHttpClient, groupId, listName).GetAwaiter().GetResult();
            if (listId != null)
            {
                return GetDocItemDownloadUrlByListId(graphHttpClient, groupId, listId, itemId).GetAwaiter().GetResult();
            }

            return null;
        }

        private async Task<string> GetDocItemDownloadUrlByListId(HttpClient graphHttpClient, string groupId,
            string listId, string itemId)
        {
            HttpResponseMessage response = await graphHttpClient.GetAsync(
                $"{Settings.GraphBaseUri}/groups/{groupId}/sites/root/lists/{listId}/items/{itemId}/driveItem");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"Get list drive item by id failed for {groupId}->{listId}->{itemId}: {response.StatusCode}-{responseMsg}");
            }

            return JObject.Parse(responseMsg)["@microsoft.graph.downloadUrl"].ToString();
        }

        private async Task<List<JObject>> GetListItemsById(HttpClient graphHttpClient, string groupId,
            string listId)
        {
            HttpResponseMessage response =
                await graphHttpClient.GetAsync(
                    $"{Settings.GraphBaseUri}/groups/{groupId}/sites/root/lists/{listId}/items?expand=fields");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"Get list items by id failed for {groupId}->{listId}: {response.StatusCode}-{responseMsg}");
            }

            GraphDataSet<JObject> dataSet = JsonConvert.DeserializeObject<GraphDataSet<JObject>>(responseMsg);
            return dataSet.Value;
        }

        private async Task<string> GetListId(HttpClient graphHttpClient, string groupId, string name)
        {
            HttpResponseMessage response =
                await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/groups/{groupId}/sites/root/lists");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new AutoTeamsStructureException($"Get list in group failed for {groupId}: {response.StatusCode}-{responseMsg}");
            }

            GraphDataSet<JObject> dataSet = JsonConvert.DeserializeObject<GraphDataSet<JObject>>(responseMsg);
            JObject matchedObject = dataSet.Value.FirstOrDefault(x =>
                x["name"].ToString().Equals(name, StringComparison.InvariantCultureIgnoreCase));
            if (matchedObject != null)
            {
                return matchedObject["id"].ToString();
            }
            else
            {
                return null;
            }
        }
    }
}