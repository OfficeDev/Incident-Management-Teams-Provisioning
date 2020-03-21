// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class GraphClientManager : IDisposable
    {
        private static AuthenticationResult authToken = null;
        private static HttpClient client = null;

        public HttpClient GetGraphHttpClient()
        {
            if (client == null)
            {
                client = new HttpClient();
            }

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GetAuthTokenAsync().Result);
            return client;
        }

        public void Dispose()
        {
            client?.Dispose();
        }
        private static async Task<string> GetAuthTokenAsync()
        {
            if (authToken == null || authToken.ExpiresOn < DateTimeOffset.UtcNow)
            {
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(Settings.ClientId)
                .WithClientSecret(Settings.ClientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{Settings.TenantId}"))
                .Build();
                string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
                authToken = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            }

            return authToken.AccessToken;
        }
    }
}
