// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Interop.AutoTeamsStructure.Graph;
using Microsoft.Interop.AutoTeamsStructure.Interfaces;

namespace Microsoft.Interop.AutoTeamsStructure
{
    public class Executor
    {
        private static readonly string DefaultChannelName = "General";
        public static void Run()
        {
            using (GraphClientManager manager = new GraphClientManager())
            {
                Factory factory = new Factory(manager);
                ITrigger trigger = factory.GetTrigger();
                ITeamsStructureExtractor teamsStructureExtractor = factory.GeTeamsStructureExtractor();
                TeamsController teamsController = new TeamsController();
                IEnumerable<string> channels = teamsStructureExtractor.GetChannels();
                Thread.Sleep(100);

                IDictionary<string, IEnumerable<FileInfo>> channelDocDictionary =
                    teamsStructureExtractor.GetChannelDocumentsDictionary();
                Thread.Sleep(100);

                IDictionary<string, IEnumerable<TeamsApp>> channelAppDictionary =
                    teamsStructureExtractor.GetChannelAppsDictionary();
                Thread.Sleep(100);

                foreach (string groupId in trigger.GetNewTeamsGroupIds())
                {
                    try
                    {
                        teamsController.CreatedTeamsFromGroupidAsync(groupId, manager.GetGraphHttpClient()).GetAwaiter().GetResult();
                        Thread.Sleep(100);

                        IEnumerable<AADIdentity> existedChannels =
                            teamsController.GetChannels(groupId, manager.GetGraphHttpClient()).GetAwaiter().GetResult();
                        Thread.Sleep(100);

                        foreach (string channelName in channels)
                        {
                            if (string.IsNullOrWhiteSpace(channelName))
                            {
                                continue;
                            }

                            string channelId = string.Empty;

                            if (channelName.Equals(DefaultChannelName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                channelId = existedChannels.FirstOrDefault(x =>
                                        x.DisplayName.Equals(DefaultChannelName,
                                            StringComparison.InvariantCultureIgnoreCase))?.Id;
                            }
                            else
                            {
                                channelId =
                                    teamsController.CreateChannelAsync(groupId, channelName, manager.GetGraphHttpClient())
                                        .GetAwaiter().GetResult();
                                Thread.Sleep(100);
                            }

                            if (channelDocDictionary.ContainsKey(channelName))
                            {
                                foreach (FileInfo file in channelDocDictionary[channelName])
                                {
                                    teamsController.UploadFileAsync(groupId, channelName,
                                        File.ReadAllBytes(file.FullName), file.Name,
                                        manager.GetGraphHttpClient()).GetAwaiter().GetResult();
                                    Thread.Sleep(100);
                                }
                            }

                            if (channelAppDictionary.ContainsKey(channelName))
                            {
                                foreach (TeamsApp app in channelAppDictionary[channelName])
                                {
                                    teamsController
                                        .AddCustomTabAsync(groupId, channelId, app, manager.GetGraphHttpClient())
                                        .GetAwaiter().GetResult();
                                    Thread.Sleep(100);
                                }
                            }
                        }

                    }
                    catch (AutoTeamsStructureException e)
                    {
                        if (e.Message.Contains("Team already exists"))
                        {
                            Console.WriteLine($"Teams team for group {groupId} already created, skip.");
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
        }
    }
}
