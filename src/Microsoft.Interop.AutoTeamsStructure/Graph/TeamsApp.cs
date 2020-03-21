// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Newtonsoft.Json;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class TeamsApp
    {
        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName { get; set; }

        [JsonProperty(PropertyName = "teamsApp@odata.bind")]
        public string TeamsAppId => $"https://graph.microsoft.com/{Settings.MsGraphApiVersion}/appCatalogs/teamsApps/{Id}";

        public string Id { get; set; }

        [JsonProperty(PropertyName = "configuration")]
        public TeamsAppConfiguration Configuration { get; set; }
    }
}
