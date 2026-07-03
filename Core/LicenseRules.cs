using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LicenceValidator.Core
{
    public sealed class Ruleset
    {
        public List<LicenseNormalizationRule> LicenseNormalization { get; set; } = new List<LicenseNormalizationRule>();
        public List<RecommendationRule> RecommendationRules { get; set; } = new List<RecommendationRule>();
        public List<UsageTableProfile> UsageTableProfiles { get; set; } = new List<UsageTableProfile>();
        public List<string> UsageExcludeEntityPatterns { get; set; } = new List<string>();
        public List<string> UsageIncludeEntityPatterns { get; set; } = new List<string>();

        private static readonly JsonSerializerOptions Relaxed = CreateRelaxedOptions();

        private static JsonSerializerOptions CreateRelaxedOptions()
        {
            return new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
        }

        public static Ruleset Load(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Rules file not found: " + path);

            return LoadFromJson(File.ReadAllText(path));
        }

        public static Ruleset LoadFromJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                throw new ArgumentException("Rules JSON is empty.");

            var rules = JsonSerializer.Deserialize<Ruleset>(json, Relaxed)
                        ?? throw new InvalidOperationException("Could not deserialize license rules.");

            if (rules.UsageTableProfiles == null) rules.UsageTableProfiles = new List<UsageTableProfile>();
            if (rules.UsageExcludeEntityPatterns == null) rules.UsageExcludeEntityPatterns = new List<string>();
            if (rules.UsageIncludeEntityPatterns == null) rules.UsageIncludeEntityPatterns = new List<string>();
            if (rules.RecommendationRules == null) rules.RecommendationRules = new List<RecommendationRule>();
            if (rules.LicenseNormalization == null) rules.LicenseNormalization = new List<LicenseNormalizationRule>();
            return rules;
        }

        public List<string> NormalizeAssignedSkus(IEnumerable<string> actualSkuPartNumbers)
        {
            var normalized = new List<string>();
            foreach (var sku in actualSkuPartNumbers.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                var match = LicenseNormalization.FirstOrDefault(rule => RegexHelper.IsMatch(sku, rule.Pattern));
                if (match != null)
                    normalized.Add(match.Normalized);
                else if (BuiltInNormalization.TryGetValue(sku, out var builtIn))
                    normalized.Add(builtIn);
                else
                    normalized.Add("Unmapped:" + sku);
            }
            return normalized
                .Where(x => !x.StartsWith("Ignore:", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static readonly Dictionary<string, string> BuiltInNormalization = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Dynamics 365 base licenses
            { "DYN365_ENTERPRISE_SALES", "SalesEnterprise" },
            { "D365_SALES_ENT", "SalesEnterprise" },
            { "D365_SALES_PRO", "SalesProfessional" },
            { "DYN365_CUSTOMER_SERVICE_PRO", "CustomerServiceProfessional" },
            { "DYN365_ENTERPRISE_CUSTOMER_SERVICE", "CustomerServiceEnterprise" },
            { "D365_CUSTOMER_SERVICE_ENT", "CustomerServiceEnterprise" },
            { "DYNAMICS_365_CONTACT_CENTER", "CustomerServiceEnterprise" },
            { "DYN365_ENTERPRISE_FIELD_SERVICE", "FieldService" },
            { "D365_FIELD_SERVICE", "FieldService" },
            { "DYN365_ENTERPRISE_TEAM_MEMBERS", "TeamMembers" },
            { "D365_TEAM_MEMBERS", "TeamMembers" },
            { "DYN365_TEAM_MEMBERS", "TeamMembers" },
            { "Dynamics_365_for_Operations_Team_members", "TeamMembers" },
            // Attach licenses
            { "D365_SALES_ENT_ATTACH", "SalesAttach" },
            { "D365_CUSTOMER_SERVICE_ENT_ATTACH", "CustomerServiceAttach" },
            { "D365_FIELD_SERVICE_ATTACH", "FieldServiceAttach" },
            // Non-D365 licenses (ignore for coverage evaluation)
            { "D365_MARKETING_USER", "Ignore:Marketing" },
            { "D365_MARKETING_APP", "Ignore:Marketing" },
            { "VIRTUAL_AGENT_USL", "Ignore:VirtualAgent" },
            { "POWERAUTOMATE_ATTENDED_RPA", "Ignore:PowerAutomateRPA" },
            { "POWERAUTOMATE_UNATTENDED_RPA", "Ignore:PowerAutomateRPA" },
            { "Microsoft_365_Copilot", "Ignore:M365Copilot" },
            { "POWER_BI_PRO", "Ignore:PowerBI" },
            { "POWER_BI_STANDARD", "Ignore:PowerBI" },
            { "PBI_PREMIUM_PER_USER", "Ignore:PowerBI" },
            { "ENTERPRISEPACK", "Ignore:Office365E3" },
            { "ENTERPRISEPREMIUM", "Ignore:Office365E5" },
            { "SPE_E3", "Ignore:M365E3" },
            { "SPE_E5", "Ignore:M365E5" },
            { "SPE_F1", "Ignore:M365F1" },
            { "FLOW_FREE", "Ignore:PowerAutomateFree" },
            { "POWERAPPS_VIRAL", "Ignore:PowerAppsViral" },
            { "POWERAPPS_DEV", "Ignore:PowerAppsDev" },
            { "POWER_APPS_PER_USER", "Ignore:PowerApps" },
            { "PROJECT_P1", "Ignore:Project" },
            { "PROJECTPREMIUM", "Ignore:Project" },
            { "VISIOCLIENT", "Ignore:Visio" },
            // Exchange / M365 add-ons
            { "EXCHANGESTANDARD", "Ignore:ExchangeOnlineP1" },
            { "EXCHANGEENTERPRISE", "Ignore:ExchangeOnlineP2" },
            { "EXCHANGE_S_ESSENTIALS", "Ignore:ExchangeEssentials" },
            // Copilot / AI add-ons
            { "CCIBOTS_PRIVPREV_VIRAL", "Ignore:CopilotStudioTrial" },
            { "COPILOT_STUDIO_VIRAL", "Ignore:CopilotStudioTrial" },
            // Teams / misc
            { "TEAMS_ESSENTIALS", "Ignore:TeamsEssentials" },
            { "MCOCAP", "Ignore:TeamsPhone" },
            { "MCOEV", "Ignore:TeamsPhoneStandard" },
            { "RIGHTSMANAGEMENT", "Ignore:AzureRMS" },
            { "INTUNE_A", "Ignore:Intune" },
            { "AAD_PREMIUM", "Ignore:EntraIDP1" },
            { "AAD_PREMIUM_P2", "Ignore:EntraIDP2" },
        };

        // Fallback lookup from well-known Microsoft license SKU GUIDs to their SkuPartNumbers.
        // Used when the tenant's subscribedSkus API call fails (e.g. missing Directory.Read.All
        // permission), so that assigned licenses can still be identified and normalized.
        public static readonly IReadOnlyDictionary<Guid, string> BuiltInSkuIdLookup = new Dictionary<Guid, string>
        {
            // Dynamics 365 Sales
            { new Guid("1e1a282c-9c54-43a2-9310-98ef728faace"), "DYN365_ENTERPRISE_SALES" },
            { new Guid("2edaa1dc-966d-4475-93d6-8ee8dfd839f4"), "D365_SALES_ENT" },
            { new Guid("c9700d56-6b8b-4d25-a518-c1c7bfd60f43"), "D365_SALES_PRO" },
            // Dynamics 365 Customer Service
            { new Guid("749742bf-0d37-4158-a120-33567104deeb"), "DYN365_ENTERPRISE_CUSTOMER_SERVICE" },
            { new Guid("1439b6e2-5d59-4873-8c59-d60e2a196e92"), "DYN365_CUSTOMER_SERVICE_PRO" },
            { new Guid("eb18b715-ea9d-4290-9994-2ebf4b5042d2"), "D365_CUSTOMER_SERVICE_ENT_ATTACH" },
            // Dynamics 365 Field Service
            { new Guid("c7d15985-e746-4f01-b113-20b575898250"), "DYN365_ENTERPRISE_FIELD_SERVICE" },
            { new Guid("a36cdaa2-a806-4b6e-9ae0-28dbd993c20e"), "D365_FIELD_SERVICE_ATTACH" },
            // Dynamics 365 Team Members
            { new Guid("6a4a1628-9b9a-424d-bed5-4118f0ede3fd"), "DYN365_ENTERPRISE_TEAM_MEMBERS" },
            // Dynamics 365 Contact Center
            { new Guid("39a78eb6-3a8a-4e1e-878a-575a5c8984e7"), "DYNAMICS_365_CONTACT_CENTER" },
            // Power Apps per-user
            { new Guid("b30411f5-fea1-4a59-9ad9-3db7c7ead579"), "POWER_APPS_PER_USER" },
            // Microsoft 365 / Office 365
            { new Guid("05e9a617-0261-4cee-bb44-138d3ef5d965"), "SPE_E3" },
            { new Guid("06ebc4ee-1bb5-47dd-8120-11324bc54e06"), "SPE_E5" },
            { new Guid("6fd2c87f-b296-42f0-b197-1e91e994b900"), "ENTERPRISEPREMIUM" },
            { new Guid("18181a46-0d4e-45cd-891e-60aabd171b4e"), "ENTERPRISEPACK" },
            { new Guid("4b9405b0-7788-4568-add1-99614e613b69"), "EXCHANGESTANDARD" },
            { new Guid("19ec0d23-8335-4cbd-94ac-6050e30712fa"), "EXCHANGEENTERPRISE" },
        };

        public RecommendationDecision EvaluateRights(UserEvidence evidence)
        {
            var matchedRules = RecommendationRules
                .Where(rule => rule.IsMatch(evidence))
                .OrderByDescending(rule => rule.Priority)
                .ThenBy(rule => rule.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var capabilities = matchedRules.Select(x => x.Capability)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var summary = capabilities.Count == 0
                ? "No confident rights-based workload match."
                : "Rights-based capability recommendation from matched roles/apps/privileges.";

            return capabilities.Count == 0
                ? RecommendationDecision.CreateReview("Rights", summary, matchedRules.Select(x => x.Name))
                : RecommendationFormatter.FromCapabilities("Rights", capabilities, matchedRules.Select(x => x.Name), summary);
        }

        public RecommendationDecision EvaluateUsage(UserUsageSummary usage)
        {
            if (!usage.HasAnyBusinessDataSignal)
                return RecommendationDecision.CreateNoSignal("Usage", "No business-data signal in the configured ownership/create/modify checks.");

            var capabilities = usage.DetectedCapabilities
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            if (capabilities.Count == 0)
            {
                var signalLabels = usage.TableSignals
                    .OrderByDescending(x => x.OwnedCount + x.CreatedCount + x.ModifiedCount)
                    .ThenBy(x => x.TableDisplayName, StringComparer.OrdinalIgnoreCase)
                    .Take(10).Select(x => x.TableDisplayName).ToList();
                return RecommendationDecision.CreateReview("Usage", "Business-data signal found, but no mapped workload capability was detected.", signalLabels);
            }

            var matchedSignals = usage.TableSignals
                .Where(x => !string.IsNullOrWhiteSpace(x.Capability))
                .OrderByDescending(x => x.OwnedCount + x.CreatedCount + x.ModifiedCount)
                .Select(x => x.TableDisplayName + ":" + (x.OwnedCount + x.CreatedCount + x.ModifiedCount)).ToList();

            return RecommendationFormatter.FromCapabilities("Usage", capabilities, matchedSignals,
                "Usage-based capability recommendation from owned/created/modified records.");
        }
    }
}