using System;
using System.Collections.Generic;

namespace LicenceValidator
{
    public class Settings
    {
        public string LastUsedOrganizationWebappUrl { get; set; }

        /// <summary>Name of the currently selected profile.</summary>
        public string ActiveProfileName { get; set; } = "Default";

        /// <summary>Named profiles for Graph credentials and audit settings.</summary>
        public List<SettingsProfile> Profiles { get; set; } = new List<SettingsProfile>();

        /// <summary>Embedded default ruleset JSON (editable in the tool).</summary>
        public string EmbeddedRulesetJson { get; set; }

        public SettingsProfile GetActiveProfile()
        {
            var name = ActiveProfileName ?? "Default";
            var match = Profiles.Find(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                match = new SettingsProfile { Name = name };
                Profiles.Add(match);
            }
            return match;
        }
    }

    public class SettingsProfile
    {
        public string Name { get; set; } = "Default";

        // Graph API credentials
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }

        // Audit settings
        public string RulesPath { get; set; }
        public string AuditMode { get; set; } = "RightsAndUsage";
        public bool IncludeDisabledUsers { get; set; }
        public int MaxDegreeOfParallelism { get; set; } = 6;
        public int UserLimit { get; set; }
        public int MaxRetryCount { get; set; } = 4;

        // Graph behavior
        public bool GraphOptional { get; set; } = true;
        public bool GraphFailOnError { get; set; }
        public string RecommendationModeWithGraph { get; set; } = "RightsThenUsage";
        public string RecommendationModeWithoutGraph { get; set; } = "RightsThenUsage";

        // Usage settings
        public bool UsageEnabled { get; set; } = true;
        public int ActivityLookbackDays { get; set; } = 180;
        public int OwnershipHistoryDays { get; set; } = 1825;
        public int BucketDays { get; set; } = 90;
        public bool AutoDiscoverCustomUserOwnedTables { get; set; } = true;
        public bool IncludeStandardUserOwnedTables { get; set; }
        public int MaxAutoDiscoveredTables { get; set; }

        public override string ToString() => Name ?? "Default";
    }
}