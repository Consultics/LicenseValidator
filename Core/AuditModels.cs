using System;
using System.Collections.Generic;
using System.Linq;

namespace LicenceValidator.Core
{
    public sealed class AuditResult
    {
        public AuditMetadata Metadata { get; set; } = new AuditMetadata();
        public List<TenantSkuRecord> TenantSkus { get; set; } = new List<TenantSkuRecord>();
        public List<AppModuleRecord> AppModules { get; set; } = new List<AppModuleRecord>();
        public List<AppModuleRoleLink> AppModuleRoles { get; set; } = new List<AppModuleRoleLink>();
        public List<UsageTableProfile> UsageTables { get; set; } = new List<UsageTableProfile>();
        public List<UserAuditResult> UserAudits { get; set; } = new List<UserAuditResult>();
    }

    public sealed class AuditMetadata
    {
        public DateTime GeneratedUtc { get; set; }
        public string DataverseUrl { get; set; } = string.Empty;
        public string RulesPath { get; set; } = string.Empty;
        public string AuditMode { get; set; } = string.Empty;
        public string RecommendationModeUsed { get; set; } = string.Empty;
        public bool GraphEnabled { get; set; }
        public bool IncludeDisabledUsers { get; set; }
        public bool UsageEnabled { get; set; }
        public int ActivityLookbackDays { get; set; }
        public int OwnershipHistoryDays { get; set; }
        public int UserCount { get; set; }
        public int HumanUserCount { get; set; }
        public int SpecialAccountCount { get; set; }
        public int KnownLicenseStateCount { get; set; }
        public int UnknownLicenseStateCount { get; set; }
    }

    public sealed class SystemUserRecord
    {
        public Guid SystemUserId { get; set; }
        public string FullName { get; set; }
        public string InternalEmailAddress { get; set; }
        public string DomainName { get; set; }
        public Guid? AzureActiveDirectoryObjectId { get; set; }
        public Guid? ApplicationId { get; set; }
        public bool IsDisabled { get; set; }
        public int? AccessMode { get; set; }
        public string UserType { get; set; } = "Human";
    }

    public sealed class EffectiveRoleRecord
    {
        public Guid RoleId { get; set; }
        public Guid? ParentRootRoleId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Source { get; set; } = "Direct";
        public Guid? SourceTeamId { get; set; }
        public string SourceTeamName { get; set; }
        public int? SourceTeamMembershipType { get; set; }
        public Guid? SourceTeamAzureActiveDirectoryObjectId { get; set; }
    }

    public sealed class TeamRecord
    {
        public Guid TeamId { get; set; }
        public string Name { get; set; } = string.Empty;
        public Guid? AzureActiveDirectoryObjectId { get; set; }
        public int? MembershipType { get; set; }
    }

    public sealed class RolePrivilegeInfo
    {
        public string PrivilegeName { get; set; } = string.Empty;
        public string Depth { get; set; } = string.Empty;
        public Guid PrivilegeId { get; set; }
        public Guid? BusinessUnitId { get; set; }
    }

    public sealed class AppModuleRecord
    {
        public Guid AppModuleId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string UniqueName { get; set; } = string.Empty;
        public bool? IsDefault { get; set; }
    }

    public sealed class AppModuleRoleLink
    {
        public Guid AppModuleId { get; set; }
        public Guid RoleId { get; set; }
    }

    public sealed class TenantSkuRecord
    {
        public Guid SkuId { get; set; }
        public string SkuPartNumber { get; set; } = string.Empty;
        public int? ConsumedUnits { get; set; }
        public string CapabilityStatus { get; set; }
        public int? EnabledUnits { get; set; }
        public int? SuspendedUnits { get; set; }
        public int? WarningUnits { get; set; }
    }

    public sealed class UserGraphLicenseSnapshot
    {
        public bool EnabledByMode { get; set; }
        public bool Attempted { get; set; }
        public bool Found { get; set; }
        public string GraphStatus { get; set; } = "Unknown";
        public string ActualLicenseState { get; set; } = "Unknown";
        public string ActualLicenseMessage { get; set; } = "Ist-Zustand konnte nicht geholt werden";
        public string UserPrincipalName { get; set; }
        public string DisplayName { get; set; }
        public bool? AccountEnabled { get; set; }
        public List<string> ActualSkuPartNumbers { get; set; } = new List<string>();
        public Dictionary<string, string> AssignmentModeBySkuPartNumber { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> Errors { get; set; } = new List<string>();

        public static UserGraphLicenseSnapshot CreateUnknown(string status, string message, bool enabledByMode = false)
        {
            return new UserGraphLicenseSnapshot { EnabledByMode = enabledByMode, GraphStatus = status, ActualLicenseState = "Unknown", ActualLicenseMessage = message };
        }
    }

    public sealed class RecommendationDecision
    {
        public string Source { get; set; } = string.Empty;
        public List<string> Capabilities { get; set; } = new List<string>();
        public List<string> MatchedRuleNames { get; set; } = new List<string>();
        public string RecommendedSku { get; set; } = "Needs review";
        public string CommercialPattern { get; set; } = "Needs review";
        public string Summary { get; set; } = string.Empty;
        public int Weight { get; set; }
        public bool IsNoSignal { get; set; }
        public bool IsReviewOnly { get; set; }

        public RecommendationDecision CloneAs(string source) => new RecommendationDecision
        {
            Source = source, Capabilities = Capabilities.ToList(), MatchedRuleNames = MatchedRuleNames.ToList(),
            RecommendedSku = RecommendedSku, CommercialPattern = CommercialPattern, Summary = Summary,
            Weight = Weight, IsNoSignal = IsNoSignal, IsReviewOnly = IsReviewOnly
        };

        public static RecommendationDecision CreateNoSignal(string source, string summary) => new RecommendationDecision
        {
            Source = source, RecommendedSku = "No usage signal / review", CommercialPattern = "No usage signal / review", Summary = summary, IsNoSignal = true
        };

        public static RecommendationDecision CreateReview(string source, string summary, IEnumerable<string> ruleNames = null) => new RecommendationDecision
        {
            Source = source, Capabilities = new List<string> { "NeedsReview" },
            MatchedRuleNames = ruleNames != null ? ruleNames.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList() : new List<string>(),
            RecommendedSku = "Needs review", CommercialPattern = "Needs review", Summary = summary,
            Weight = RecommendationFormatter.CapabilityRank("NeedsReview"), IsReviewOnly = true
        };
    }

    public sealed class UserUsageSummary
    {
        public long OwnedRecordCount { get; set; }
        public long CreatedRecordCount { get; set; }
        public long ModifiedRecordCount { get; set; }
        public long BusinessSignalRecordCount { get; set; }
        public Dictionary<string, long> CapabilitySignalCounts { get; set; } = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        public List<string> DetectedCapabilities { get; set; } = new List<string>();
        public List<UserUsageTableSignal> TableSignals { get; set; } = new List<UserUsageTableSignal>();
        public string Status { get; set; } = string.Empty;
        public List<string> Notes { get; set; } = new List<string>();
        public bool HasAnyBusinessDataSignal => BusinessSignalRecordCount > 0;
        public bool HasRecentCreateOrModify => CreatedRecordCount > 0 || ModifiedRecordCount > 0;
    }

    public sealed class UserUsageTableSignal
    {
        public string TableLogicalName { get; set; } = string.Empty;
        public string TableDisplayName { get; set; } = string.Empty;
        public string Capability { get; set; } = string.Empty;
        public long OwnedCount { get; set; }
        public long CreatedCount { get; set; }
        public long ModifiedCount { get; set; }
    }

    public sealed class UserAuditResult
    {
        public SystemUserRecord User { get; set; } = new SystemUserRecord();
        public List<EffectiveRoleRecord> EffectiveRoles { get; set; } = new List<EffectiveRoleRecord>();
        public List<TeamRecord> EffectiveTeams { get; set; } = new List<TeamRecord>();
        public List<RolePrivilegeInfo> EffectivePrivileges { get; set; } = new List<RolePrivilegeInfo>();
        public List<AppModuleRecord> AccessibleApps { get; set; } = new List<AppModuleRecord>();
        public UserGraphLicenseSnapshot Graph { get; set; } = new UserGraphLicenseSnapshot();
        public List<string> ActualAssignedNormalized { get; set; } = new List<string>();
        public RecommendationDecision RightsRecommendation { get; set; } = RecommendationDecision.CreateReview("Rights", "Not evaluated yet.");
        public RecommendationDecision UsageRecommendation { get; set; } = RecommendationDecision.CreateNoSignal("Usage", "Usage not evaluated.");
        public RecommendationDecision FinalRecommendation { get; set; } = RecommendationDecision.CreateReview("Final", "Not evaluated yet.");
        public UserUsageSummary Usage { get; set; } = new UserUsageSummary();
        public string CoverageStatus { get; set; } = string.Empty;
        public string OverallStatus { get; set; } = string.Empty;
        public string OptimizationStatus { get; set; } = string.Empty;
        public string SuggestedAction { get; set; } = string.Empty;
        public List<string> Notes { get; set; } = new List<string>();
    }

    public sealed class UserEvidence
    {
        public List<string> RoleNames { get; set; } = new List<string>();
        public List<string> AppNames { get; set; } = new List<string>();
        public List<string> PrivilegeNames { get; set; } = new List<string>();
        public List<string> AssignedSkuPartNumbers { get; set; } = new List<string>();
        public bool IsDisabled { get; set; }
        public int? AccessMode { get; set; }
        public string UserType { get; set; }
    }

    public sealed class CoverageResult
    {
        public string Status { get; set; } = string.Empty;
        public List<string> Notes { get; set; } = new List<string>();
    }

    public sealed class UsageTableProfile
    {
        public string LogicalName { get; set; } = string.Empty;
        public string EntitySetName { get; set; } = string.Empty;
        public string PrimaryIdAttribute { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Capability { get; set; } = string.Empty;
        public bool CountOwned { get; set; } = true;
        public bool CountCreated { get; set; } = true;
        public bool CountModified { get; set; } = true;
        public bool CountsAsBusinessSignal { get; set; } = true;
        public bool IsAutoDiscovered { get; set; }
        public bool IsCustomEntity { get; set; }
        public bool IsActivity { get; set; }
        public string QueryError { get; set; }
    }

    public sealed class DataverseEntityMetadata
    {
        public string LogicalName { get; set; } = string.Empty;
        public string SchemaName { get; set; } = string.Empty;
        public string EntitySetName { get; set; } = string.Empty;
        public string PrimaryIdAttribute { get; set; } = string.Empty;
        public bool IsCustomEntity { get; set; }
        public bool IsActivity { get; set; }
    }

    public sealed class LicenseNormalizationRule
    {
        public string Pattern { get; set; } = string.Empty;
        public string Normalized { get; set; } = string.Empty;
    }

    public sealed class RecommendationRule
    {
        public string Name { get; set; } = string.Empty;
        public int Priority { get; set; }
        public string Capability { get; set; } = string.Empty;
        public string Mode { get; set; } = "any";
        public List<string> AnyRolePatterns { get; set; } = new List<string>();
        public List<string> AnyAppPatterns { get; set; } = new List<string>();
        public List<string> AnyPrivilegePatterns { get; set; } = new List<string>();
        public List<string> AnyAssignedSkuPatterns { get; set; } = new List<string>();
        public List<string> ExcludeRolePatterns { get; set; } = new List<string>();
        public List<string> ExcludeAppPatterns { get; set; } = new List<string>();
        public List<string> ExcludePrivilegePatterns { get; set; } = new List<string>();

        public bool IsMatch(UserEvidence evidence)
        {
            if (HasAnyMatch(evidence.RoleNames, ExcludeRolePatterns) || HasAnyMatch(evidence.AppNames, ExcludeAppPatterns) || HasAnyMatch(evidence.PrivilegeNames, ExcludePrivilegePatterns)) return false;
            var checks = new List<bool>();
            if (AnyRolePatterns.Count > 0) checks.Add(HasAnyMatch(evidence.RoleNames, AnyRolePatterns));
            if (AnyAppPatterns.Count > 0) checks.Add(HasAnyMatch(evidence.AppNames, AnyAppPatterns));
            if (AnyPrivilegePatterns.Count > 0) checks.Add(HasAnyMatch(evidence.PrivilegeNames, AnyPrivilegePatterns));
            if (AnyAssignedSkuPatterns.Count > 0) checks.Add(HasAnyMatch(evidence.AssignedSkuPartNumbers, AnyAssignedSkuPatterns));
            if (checks.Count == 0) return false;
            return string.Equals(Mode, "all", StringComparison.OrdinalIgnoreCase) ? checks.All(x => x) : checks.Any(x => x);
        }

        private static bool HasAnyMatch(IEnumerable<string> values, IEnumerable<string> patterns)
        {
            var vl = values.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            var pl = patterns.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            if (vl.Count == 0 || pl.Count == 0) return false;
            foreach (var v in vl) foreach (var p in pl) if (RegexHelper.IsMatch(v, p)) return true;
            return false;
        }
    }
}