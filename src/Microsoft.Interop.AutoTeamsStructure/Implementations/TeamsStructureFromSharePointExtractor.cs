// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Microsoft.Interop.AutoTeamsStructure.Graph;
using Microsoft.Interop.AutoTeamsStructure.Interfaces;
using Newtonsoft.Json.Linq;

namespace Microsoft.Interop.AutoTeamsStructure.Implementations
{
    public class TeamsStructureFromSharePointExtractor : ITeamsStructureExtractor
    {
        private const string TempFileFolder = "./Temp";
        private readonly SharePointController sharePointController;
        private readonly GraphClientManager graphClientManager;

        public TeamsStructureFromSharePointExtractor(SharePointController _controller, GraphClientManager _graphClientManager)
        {
            sharePointController = _controller ?? throw new ArgumentException(nameof(_controller));
            graphClientManager = _graphClientManager ?? throw new ArgumentException(nameof(_graphClientManager));
        }

        public IEnumerable<string> GetChannels()
        {
            return GetChannelListItems(Settings.SharePointGroupId, Settings.ChannelListName)
                .Select(x => x["fields"]?[Settings.ChannelNameField]?.ToString()).Distinct();
        }

        public IDictionary<string, IEnumerable<FileInfo>> GetChannelDocumentsDictionary()
        {
            Dictionary<string, IEnumerable<FileInfo>> result = new Dictionary<string, IEnumerable<FileInfo>>();
            IEnumerable<JObject> channels = GetChannelListItems(Settings.SharePointGroupId, Settings.ChannelListName);
            IEnumerable<JObject> teamsChannelDoclistItems = sharePointController.GetListItems(graphClientManager.GetGraphHttpClient(),
                Settings.SharePointGroupId, Settings.ChannelDocListName);
            foreach (JObject item in teamsChannelDoclistItems)
            {
                string channelid = item["fields"]?[Settings.ChannelDocChannelNameField]?.ToString();
                if (channelid != null)
                {
                    string channelName =
                        channels.FirstOrDefault(x => x["fields"]?["id"]?.ToString() == channelid)["fields"]?
                        [Settings.ChannelNameField]?.ToString();
                    if (!string.IsNullOrWhiteSpace(channelName))
                    {
                        string fileName = item["fields"]?[Settings.ChannelDocFileNameField]?.ToString();
                        if (!string.IsNullOrWhiteSpace(fileName))
                        {
                            string fileDownloadUrl = sharePointController.GetDocItemDownloadUrl(graphClientManager.GetGraphHttpClient(),
                                Settings.SharePointGroupId, Settings.ChannelDocListName, item["id"].ToString());
                            string tempFileName = $"{TempFileFolder}/{fileName}";
                            DownLoadTemplateFile(tempFileName, fileDownloadUrl);
                            if (result.ContainsKey(channelName))
                            {
                                ((List<FileInfo>)result[channelName]).Add(new FileInfo(tempFileName));
                            }
                            else
                            {
                                List<FileInfo> fileList = new List<FileInfo>
                                {
                                    new FileInfo(tempFileName)
                                };
                                result.Add(channelName, fileList);
                            }
                        }
                        else
                        {
                            throw new AutoTeamsStructureException($"Document file name cannot be found in document template list for item {item["id"]}");
                        }
                    }
                    else
                    {
                        throw new AutoTeamsStructureException($"Channel name cannot be found in document template list for item {item["id"]}");
                    }
                }
            }

            return result;
        }

        public IDictionary<string, IEnumerable<TeamsApp>> GetChannelAppsDictionary()
        {
            Dictionary<string, IEnumerable<TeamsApp>> result = new Dictionary<string, IEnumerable<TeamsApp>>();
            IEnumerable<JObject> channels = GetChannelListItems(Settings.SharePointGroupId, Settings.ChannelListName);
            IEnumerable<JObject> teamsChannelApplistItems = sharePointController.GetListItems(graphClientManager.GetGraphHttpClient(),
                Settings.SharePointGroupId, Settings.ChannelAppListName);
            foreach (JObject item in teamsChannelApplistItems)
            {
                string channelid = item["fields"]?[Settings.ChannelAppChannelNameField]?.ToString();
                if (channelid != null)
                {
                    string channelName =
                        channels.FirstOrDefault(x => x["fields"]?["id"]?.ToString() == channelid)["fields"]?
                        [Settings.ChannelNameField]?.ToString();
                    if (!string.IsNullOrWhiteSpace(channelName))
                    {
                        TeamsApp app = new TeamsApp
                        {
                            DisplayName = item["fields"]?[Settings.ChannelAppTabNameField]?.ToString(),
                            Id = item["fields"]?[Settings.ChannelAppAppIdField]?.ToString()
                        };

                        TeamsAppConfiguration appConfig = new TeamsAppConfiguration
                        {
                            EntityId = item["fields"]?[Settings.ChannelAppAppEntityIdField]?.ToString(),
                            ContentUrl = item["fields"]?[Settings.ChannelAppContentUrlField]?.ToString(),
                            WebsiteUrl = item["fields"]?[Settings.ChannelAppWebUrlField]?.ToString(),
                            RemoveUrl = item["fields"]?[Settings.ChannelAppRemoveUrlField]?.ToString()
                        };

                        app.Configuration = appConfig;

                        if (result.ContainsKey(channelName))
                        {
                            ((List<TeamsApp>)result[channelName]).Add(app);
                        }
                        else
                        {
                            List<TeamsApp> appList = new List<TeamsApp>
                            {
                                app
                            };
                            result.Add(channelName, appList);
                        }
                    }
                    else
                    {
                        throw new AutoTeamsStructureException($"Channel name cannot be found in app list for item {item["id"]}");
                    }
                }
            }

            return result;
        }

        private void DownLoadTemplateFile(string fileName, string downloadUrl)
        {
            if (!Directory.Exists(TempFileFolder))
            {
                Directory.CreateDirectory(TempFileFolder);
            }

            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(
                    new Uri(downloadUrl),
                    $"{fileName}"
                );
            }
        }

        private IEnumerable<JObject> GetChannelListItems(string siteGroupId, string channelListName)
        {
            IEnumerable<JObject> listItems = sharePointController.GetListItems(graphClientManager.GetGraphHttpClient(),
                siteGroupId, channelListName);
            if (listItems == null)
            {
                throw new AutoTeamsStructureException("No items can be found in channel list");
            }

            return listItems;
        }
    }
}
