// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using Microsoft.Interop.AutoTeamsStructure.Graph;
using Microsoft.Interop.AutoTeamsStructure.Implementations;
using Microsoft.Interop.AutoTeamsStructure.Interfaces;

namespace Microsoft.Interop.AutoTeamsStructure
{
    public class Factory
    {
        private readonly GraphClientManager graphClientManager;
        public Factory(GraphClientManager _graphClientManager)
        {
            graphClientManager = _graphClientManager ?? throw new ArgumentException(nameof(_graphClientManager));
        }

        public ITrigger GetTrigger()
        {
            GroupsController groupsController = new GroupsController();
            return new Covid19DataTrigger(groupsController, graphClientManager);
        }

        public ITeamsStructureExtractor GeTeamsStructureExtractor()
        {
            SharePointController spController = new SharePointController();
            return new TeamsStructureFromSharePointExtractor(spController, graphClientManager);
        }
    }
}
