// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Microsoft.Interop.AutoTeamsStructure.Graph
{
    public class GraphClientManager : IDisposable
    {
        private static AuthenticationResult authToken = null;
        private static AuthenticationResult delegateAuthToken = null;
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

        public HttpClient GetDelegateGraphClient()
        {
            if (client == null)
            {
                client = new HttpClient();
            }

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GetDelegateAuthTokenAsync().Result);
            return client;
        }

        private async static Task<string> GetDelegateAuthTokenAsync()
        {
            if (delegateAuthToken != null && delegateAuthToken.ExpiresOn > DateTimeOffset.UtcNow)
            {
                return delegateAuthToken.AccessToken;
            }

            string authority = $"https://login.microsoftonline.com/{Settings.TenantId}";
            string[] scopes = new string[] { "openid", "Sites.ReadWrite.All" };
            IPublicClientApplication app = PublicClientApplicationBuilder.Create(Settings.ClientId)
                .WithAuthority(authority)
                .Build();
            var accounts = await app.GetAccountsAsync();
            if (accounts.Any())
            {
                delegateAuthToken = await app.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {
                try
                {
                    var securePassword = new SecureString();
                    foreach (char c in Settings.DelegatedUserPwd)
                    {
                        securePassword.AppendChar(c);
                    }

                    delegateAuthToken = await app.AcquireTokenByUsernamePassword(scopes,
                            Settings.DelegatedUserName,
                            securePassword)
                        .ExecuteAsync();
                }
                catch (MsalException)
                {
                    throw new AutoTeamsStructureException("Cannot get delegate token");
                }
            }

            return delegateAuthToken.AccessToken;
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
