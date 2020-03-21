// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.IO;
using Microsoft.Interop.AutoTeamsStructure.Graph;

namespace Microsoft.Interop.AutoTeamsStructure.Interfaces
{
    public interface ITeamsStructureExtractor
    {
        IEnumerable<string> GetChannels();

        IDictionary<string, IEnumerable<FileInfo>> GetChannelDocumentsDictionary();

        IDictionary<string, IEnumerable<TeamsApp>> GetChannelAppsDictionary();
    }
}
