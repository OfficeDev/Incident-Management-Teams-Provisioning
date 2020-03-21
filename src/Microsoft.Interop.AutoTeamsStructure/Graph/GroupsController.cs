// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
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
    }
}
