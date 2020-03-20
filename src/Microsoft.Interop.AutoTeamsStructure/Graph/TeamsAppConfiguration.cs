// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Newtonsoft.Json;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class TeamsAppConfiguration
    {
        [JsonProperty(PropertyName = "entityId")]
        public string EntityId { get; set; }

        [JsonProperty(PropertyName = "contentUrl")]
        public string ContentUrl { get; set; }

        [JsonProperty(PropertyName = "websiteUrl")]
        public string WebsiteUrl { get; set; }

        [JsonProperty(PropertyName = "removeUrl")]
        public string RemoveUrl { get; set; }
    }
}
