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
    public class NationalWeatherEventTrigger : ITrigger
    {
        private readonly GroupsController groupsController = null;
        private readonly GraphClientManager graphClientManager = null;
        private Dictionary<string, string> groupToAddtionalChannelDic = null;

        public NationalWeatherEventTrigger(GroupsController _controller, GraphClientManager _graphClientManager)
        {
            groupsController = _controller ?? throw new ArgumentException(nameof(_controller));
            graphClientManager = _graphClientManager ?? throw new ArgumentException(nameof(_graphClientManager));
        }

        public IEnumerable<string> GetTeamsGroupIds()
        {
            if (groupToAddtionalChannelDic == null)
            {
                CreateGroupIdToChannelNameDic();
            }

            return groupToAddtionalChannelDic.Keys;
        }

        public string GetCustomChannelForGroup(string groupId)
        {
            if (groupToAddtionalChannelDic == null)
            {
                CreateGroupIdToChannelNameDic();
            }

            return groupToAddtionalChannelDic.ContainsKey(groupId) ? groupToAddtionalChannelDic[groupId] : null;
        }

        private void CreateGroupIdToChannelNameDic()
        {
            groupToAddtionalChannelDic = new Dictionary<string, string>();
            Dictionary<string, string> countyToAlertDic = CreateCountyToAlertDic();
            IEnumerable<AADIdentity> groups = groupsController.GetGroups(graphClientManager.GetGraphHttpClient())
                .GetAwaiter().GetResult();
            foreach (string countyName in countyToAlertDic.Keys)
            {
                bool groupAlreayCreated = false;
                foreach (AADIdentity identity in groups)
                {
                    if (identity.DisplayName.Equals(countyName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        groupToAddtionalChannelDic.Add(identity.Id, countyToAlertDic[countyName]);
                        groupAlreayCreated = true;
                        break;
                    }
                }

                if (!groupAlreayCreated)
                {
                    AADIdentity group = groupsController
                        .CreateGroup(graphClientManager.GetGraphHttpClient(), countyName).GetAwaiter().GetResult();
                    groupToAddtionalChannelDic.Add(group.Id, countyToAlertDic[countyName]);
                }
            }
        }

        private Dictionary<string, string> CreateCountyToAlertDic()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            List<JToken> alertList = GetNationalWeatherAlerts();
            foreach (JToken alert in alertList)
            {
                foreach (var zone in alert["properties"]["affectedZones"])
                {
                    string zoneUrl = zone.ToString();
                    if (zoneUrl.Contains("/zones/county"))
                    {
                        string countyName = GetCountyName(zoneUrl);
                        string alertName = DateTime.Now.ToString("MM-dd-yyyy") + " "
                                           + alert["properties"]["event"].ToString();
                        if (!result.Keys.Contains(countyName))
                        {
                            result.Add(countyName, alertName);
                        }

                        break;
                    }
                }
            }

            return result;
        }

        private string GetCountyName(string zoneUrl)
        {
            JToken zoneToken = JToken.Parse(GetResponeContent(zoneUrl));
            return $"{zoneToken["properties"]["name"]} County, {zoneToken["properties"]["state"]}";
        }

        private List<JToken> GetNationalWeatherAlerts()
        {
            List<JToken> alertList = new List<JToken>();
            JToken token = JToken.Parse(GetResponeContent(Settings.NationalWeatherAlertApi));
            JArray features = (JArray)token["features"];
            int alertNumber = 0;
            foreach (var child in token["features"])
            {
                string messageType = child["properties"]["messageType"].ToString().Trim();
                string serverity = child["properties"]["severity"].ToString().Trim();
                string status = child["properties"]["status"].ToString().Trim();
                string affectedZones = child["properties"]["affectedZones"].ToString().Trim();
                if (messageType == "Alert" && serverity == "Severe" && status == "Actual" && affectedZones.Contains("/zones/county"))
                {
                    alertList.Add(child);
                    alertNumber++;
                    if (alertNumber == Settings.AlertsLimit)
                    {
                        break;
                    }
                }
            }

            return alertList;
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
