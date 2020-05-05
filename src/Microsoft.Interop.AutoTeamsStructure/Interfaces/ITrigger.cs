// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Generic;

namespace Microsoft.Interop.AutoTeamsStructure.Interfaces
{
    public interface ITrigger
    {
        /// <summary>
        /// The groups ids for creating Teams team
        /// </summary>
        /// <returns>Group ids</returns>
        IEnumerable<string> GetTeamsGroupIds();

        /// <summary>
        /// Get the customized channel name need to be created for specified group id
        /// </summary>
        /// <param name="groupId">The group id</param>
        /// <returns>Customized channel name need to be created</returns>
        string GetCustomChannelForGroup(string groupId);
    }
}
