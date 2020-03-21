// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.Interop.AutoTeamsStructure.Graph;
using Microsoft.Interop.AutoTeamsStructure.Interfaces;
using Newtonsoft.Json.Linq;

namespace Microsoft.Interop.AutoTeamsStructure.Implementations
{
    public class Covid19DataTrigger : ITrigger
    {
        private readonly string JohnsHopkinsDataQueryForUsPath = Settings.CovId19DataQueryUri;

        private readonly GroupsController groupsController = null;
        private readonly GraphClientManager graphClientManager = null;

        public Covid19DataTrigger(GroupsController _controller, GraphClientManager _graphClientManager)
        {
            groupsController = _controller ?? throw new ArgumentException(nameof(_controller));
            graphClientManager = _graphClientManager ?? throw new ArgumentException(nameof(_graphClientManager));
        }

        public IEnumerable<string> GetNewTeamsGroupIds()
        {
            List<string> groupIdList = new List<string>();
            IEnumerable<AADIdentity> groups = groupsController.GetGroups(graphClientManager.GetGraphHttpClient())
                .GetAwaiter().GetResult();

            foreach (string stateName in GetUSStateNamesHasConfirmedCases())
            {
                AADIdentity group = groups.FirstOrDefault(x =>
                    x.DisplayName.Equals(stateName, StringComparison.InvariantCultureIgnoreCase));
                if (group == null)
                {
                    Console.WriteLine($"Warning: Cannot find group with display name '{stateName}'");
                }
                else
                {
                    groupIdList.Add(group.Id);
                }
            }

            return groupIdList;
        }

        private List<string> GetUSStateNamesHasConfirmedCases()
        {
            List<string> stateNameList = new List<string>();
            JToken token = JToken.Parse(GetResponeContent(JohnsHopkinsDataQueryForUsPath));
            foreach (var child in token["features"])
            {
                string confirmed = child["attributes"]["Confirmed"].ToString().Trim();
                if (confirmed != "0")
                {
                    stateNameList.Add(child["attributes"]["Province_State"].ToString());
                }
            }

            return stateNameList;
        }

        private string GetResponeContent(string url)
        {
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.UserAgent = "request";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            var encoding = Encoding.UTF8;
            using (var reader = new StreamReader(response.GetResponseStream(), encoding))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
