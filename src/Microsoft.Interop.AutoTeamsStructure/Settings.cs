// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Configuration;

namespace Microsoft.Interop.AutoTeamsStructure
{
    public class Settings
    {
        private const string JohnHopkinsDataQuery = "JohnHopkinsDataQuery";
        private const string GraphApiVersion = "GraphApiVersion";
        private const string AppTenantId = "AppTenantId";
        private const string AppClientId = "AppClientId";
        private const string AppClientSecret = "AppClientSecret";
        private const string SpGroupId = "SharePointGroupId";
        private const string ChannelListNameInSp = "ChannelListNameInSP";
        private const string ChannelNameFieldInSp = "ChannelNameFieldInSp";
        private const string ChannelAppListNameInSp = "ChannelAppListNameInSP";
        private const string ChannelAppChannelNameFieldInSp = "ChannelApp_ChannelNameLookupFieldInSP";
        private const string ChannelAppTabNameFieldInSp = "ChannelApp_TabNameFieldInSP";
        private const string ChannelAppAppIdFieldInSp = "ChannelApp_AppIdFieldInSP";
        private const string ChannelAppAppEntityIdFieldInSp = "ChannelApp_AppEntityIdFieldInSP";
        private const string ChannelAppWebUrlFieldInSp = "ChannelApp_WebUrlFieldInSP";
        private const string ChannelAppContentUrlFieldInSp = "ChannelApp_ContentUrlFieldInSP";
        private const string ChannelAppRemoveUrlFieldInSp = "ChannelApp_RemoveUrlFieldInSP";
        private const string ChannelDocListNameInSp = "ChannelDocListNameInSP";
        private const string ChannelDocChannelNameFieldInSp = "ChannelDoc_ChannelNameLookupFieldInSP";
        private const string ChannelDocFileNameFieldInSp = "ChannelDoc_FileNameFieldInSP";
        private const string LookupFieldPostFixInSharePoint = "LookupId";


        public static string GraphBaseUri = $"https://graph.microsoft.com/{MsGraphApiVersion}";

        public static string CovId19DataQueryUri => ConfigurationManager.AppSettings[JohnHopkinsDataQuery];

        public static string MsGraphApiVersion => ConfigurationManager.AppSettings[GraphApiVersion];

        public static string TenantId => ConfigurationManager.AppSettings[AppTenantId];

        public static string ClientId => ConfigurationManager.AppSettings[AppClientId];

        public static string ClientSecret => ConfigurationManager.AppSettings[AppClientSecret];

        public static string SharePointGroupId => ConfigurationManager.AppSettings[SpGroupId];

        public static string ChannelListName => ConfigurationManager.AppSettings[ChannelListNameInSp];

        public static string ChannelNameField => ConfigurationManager.AppSettings[ChannelNameFieldInSp];

        public static string ChannelAppListName => ConfigurationManager.AppSettings[ChannelAppListNameInSp];

        public static string ChannelAppChannelNameField =>
            ConfigurationManager.AppSettings[ChannelAppChannelNameFieldInSp] + LookupFieldPostFixInSharePoint;

        public static string ChannelAppTabNameField => ConfigurationManager.AppSettings[ChannelAppTabNameFieldInSp];

        public static string ChannelAppAppIdField => ConfigurationManager.AppSettings[ChannelAppAppIdFieldInSp];

        public static string ChannelAppAppEntityIdField => ConfigurationManager.AppSettings[ChannelAppAppEntityIdFieldInSp];

        public static string ChannelAppWebUrlField => ConfigurationManager.AppSettings[ChannelAppWebUrlFieldInSp];

        public static string ChannelAppContentUrlField => ConfigurationManager.AppSettings[ChannelAppContentUrlFieldInSp];

        public static string ChannelAppRemoveUrlField => ConfigurationManager.AppSettings[ChannelAppRemoveUrlFieldInSp];

        public static string ChannelDocListName => ConfigurationManager.AppSettings[ChannelDocListNameInSp];

        public static string ChannelDocChannelNameField =>
            ConfigurationManager.AppSettings[ChannelDocChannelNameFieldInSp] + LookupFieldPostFixInSharePoint;

        public static string ChannelDocFileNameField => ConfigurationManager.AppSettings[ChannelDocFileNameFieldInSp];

    }
}
