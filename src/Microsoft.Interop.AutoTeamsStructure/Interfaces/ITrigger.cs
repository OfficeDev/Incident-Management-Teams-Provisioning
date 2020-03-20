// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Generic;

namespace Microsoft.Interop.AutoTeamsStructure.Interfaces
{
    public interface ITrigger
    {
        IEnumerable<string> GetNewTeamsGroupIds();
    }
}
