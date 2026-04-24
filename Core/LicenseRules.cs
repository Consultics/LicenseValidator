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

        private static readonly JsonSerializerOptions Relaxed = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

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
                normalized.Add(match != null ? match.Normalized : "Unmapped:" + sku);
            }
            return normalized
                .Where(x => !x.StartsWith("Ignore:", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

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