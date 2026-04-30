using System;
using System.Collections.Generic;

namespace LicenceValidator.Core
{
    public static class AppModes
    {
        public const string RightsOnly = "RightsOnly";
        public const string RightsAndUsage = "RightsAndUsage";
        public const string NoGraphRightsOnly = "NoGraphRightsOnly";
        public const string NoGraphRightsAndUsage = "NoGraphRightsAndUsage";

        public static bool UsesGraph(string mode) =>
            !string.Equals(mode, NoGraphRightsOnly, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(mode, NoGraphRightsAndUsage, StringComparison.OrdinalIgnoreCase);

        public static bool UsesUsage(string mode) =>
            string.Equals(mode, RightsAndUsage, StringComparison.OrdinalIgnoreCase)
            || string.Equals(mode, NoGraphRightsAndUsage, StringComparison.OrdinalIgnoreCase);
    }

    public static class RecommendationModes
    {
        public const string Rights = "Rights";
        public const string Usage = "Usage";
        public const string RightsThenUsage = "RightsThenUsage";
        public const string HigherOfRightsAndUsage = "HigherOfRightsAndUsage";
    }

    public sealed class ToolConfig
    {
        public string TenantId { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
        public string DataverseUrl { get; set; } = string.Empty;
        public string GraphBaseUrl { get; set; } = "https://graph.microsoft.com/v1.0";
        public string AuditMode { get; set; } = AppModes.RightsAndUsage;
        public bool IncludeDisabledUsers { get; set; }
        public int MaxDegreeOfParallelism { get; set; } = 6;
        public int UserLimit { get; set; }
        public int MaxRetryCount { get; set; } = 4;
        public bool GraphOptional { get; set; } = true;
        public bool GraphFailOnError { get; set; }
        public string RecommendationModeWithGraph { get; set; } = RecommendationModes.RightsThenUsage;
        public string RecommendationModeWithoutGraph { get; set; } = RecommendationModes.RightsThenUsage;
        public bool UsageEnabled { get; set; } = true;
        public int ActivityLookbackDays { get; set; } = 180;
        public int OwnershipHistoryDays { get; set; } = 1825;
        public int BucketDays { get; set; } = 90;
        public bool AutoDiscoverCustomUserOwnedTables { get; set; } = true;
        public bool IncludeStandardUserOwnedTables { get; set; }
        public int MaxAutoDiscoveredTables { get; set; }
        public List<string> IncludeEntityLogicalNamePatterns { get; set; } = new List<string>();
        public List<string> ExcludeEntityLogicalNamePatterns { get; set; } = new List<string>();

        public bool GraphEnabled => AppModes.UsesGraph(AuditMode);
        public bool UsageEnabledEffective => AppModes.UsesUsage(AuditMode) && UsageEnabled;
        public bool ShouldFailOnGraphError => GraphEnabled && GraphFailOnError && !GraphOptional;
        public string EffectiveRecommendationMode => GraphEnabled ? RecommendationModeWithGraph : RecommendationModeWithoutGraph;
    }
}