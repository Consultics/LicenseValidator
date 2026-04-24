using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace LicenceValidator.Core
{
    public static class UserClassifier
    {
        public static string DetermineUserType(SystemUserRecord user)
        {
            if (user.ApplicationId.HasValue) return "ApplicationUser";
            switch (user.AccessMode)
            {
                case 1: return "Administrative";
                case 2: return "ReadOnly";
                case 3: return "SupportUser";
                case 4: return "NonInteractive";
                case 5: return "DelegatedAdmin";
                default: return "Human";
            }
        }

        public static bool IsApplicationUser(SystemUserRecord user) =>
            user.ApplicationId.HasValue || string.Equals(user.UserType, "ApplicationUser", StringComparison.OrdinalIgnoreCase);

        public static bool IsSpecialAccount(SystemUserRecord user)
        {
            return IsApplicationUser(user)
                   || string.Equals(user.UserType, "Administrative", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(user.UserType, "SupportUser", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(user.UserType, "NonInteractive", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(user.UserType, "DelegatedAdmin", StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class RecommendationFormatter
    {
        public static RecommendationDecision FromCapabilities(string source, IReadOnlyCollection<string> capabilities, IEnumerable<string> matchedRuleNames = null, string summary = null)
        {
            var normalized = NormalizeCapabilities(capabilities);
            if (normalized.Count == 0)
                return RecommendationDecision.CreateReview(source, summary ?? "No confident workload match.", matchedRuleNames);

            return new RecommendationDecision
            {
                Source = source,
                Capabilities = normalized,
                MatchedRuleNames = matchedRuleNames != null
                    ? matchedRuleNames.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    : new List<string>(),
                RecommendedSku = ToPreferredSku(normalized),
                CommercialPattern = ToCommercialPattern(normalized),
                Summary = summary ?? ToCommercialPattern(normalized),
                Weight = normalized.Select(CapabilityRank).DefaultIfEmpty(0).Max()
            };
        }

        public static List<string> NormalizeCapabilities(IEnumerable<string> capabilities)
        {
            var list = capabilities.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (list.Count > 1 && list.Any(x => string.Equals(x, "TeamMembers", StringComparison.OrdinalIgnoreCase)))
                list.RemoveAll(x => string.Equals(x, "TeamMembers", StringComparison.OrdinalIgnoreCase));
            return list.OrderByDescending(CapabilityRank).ThenBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
        }

        public static string ToCommercialPattern(IReadOnlyCollection<string> capabilities)
        {
            if (capabilities.Count == 0) return "Needs review";
            if (capabilities.Count == 1) return CapabilityToPattern(capabilities.First());
            if (capabilities.Any(x => string.Equals(x, "NeedsReview", StringComparison.OrdinalIgnoreCase))) return "Needs review";
            return "Multiple workloads: " + string.Join(" + ", capabilities.Select(CapabilityToLabel)) + ". Validate the commercial base/attach combination manually.";
        }

        public static string ToPreferredSku(IReadOnlyCollection<string> capabilities)
        {
            if (capabilities.Count == 0) return "Needs review";
            if (capabilities.Count == 1) return CapabilityToPreferredSku(capabilities.First());
            if (capabilities.Any(x => string.Equals(x, "NeedsReview", StringComparison.OrdinalIgnoreCase))) return "Needs review";
            return "Multiple workloads (validate base/attach)";
        }

        public static string CapabilityToPattern(string cap)
        {
            switch (cap)
            {
                case "TeamMembers": return "Dynamics 365 Team Members (scenario restrictions must still be validated)";
                case "SalesFull": return "Dynamics 365 Sales Professional or Sales Enterprise/Premium";
                case "CustomerServiceFull": return "Dynamics 365 Customer Service Professional or Enterprise/Premium";
                case "FieldServiceFull": return "Dynamics 365 Field Service, or Field Service Attach plus another qualifying Dynamics 365 base";
                default: return "Needs review";
            }
        }

        public static string CapabilityToPreferredSku(string cap)
        {
            switch (cap)
            {
                case "TeamMembers": return "TeamMembers";
                case "SalesFull": return "SalesProfessional or SalesEnterprise/Premium";
                case "CustomerServiceFull": return "CustomerServiceProfessional or CustomerServiceEnterprise/Premium";
                case "FieldServiceFull": return "FieldService or FieldServiceAttach + qualifying base";
                default: return "Needs review";
            }
        }

        public static string CapabilityToLabel(string cap)
        {
            switch (cap)
            {
                case "TeamMembers": return "Team Members";
                case "SalesFull": return "Sales";
                case "CustomerServiceFull": return "Customer Service";
                case "FieldServiceFull": return "Field Service";
                case "NeedsReview": return "Needs review";
                default: return cap;
            }
        }

        public static int CapabilityRank(string cap)
        {
            switch (cap)
            {
                case "TeamMembers": return 10;
                case "SalesFull": return 30;
                case "CustomerServiceFull": return 30;
                case "FieldServiceFull": return 40;
                case "NeedsReview": return 5;
                default: return 0;
            }
        }
    }

    public static class RecommendationCombiner
    {
        public static RecommendationDecision Combine(RecommendationDecision rights, RecommendationDecision usage, string mode)
        {
            var rightsConcrete = RecommendationFormatter.NormalizeCapabilities(
                rights.Capabilities.Where(x => !string.Equals(x, "NeedsReview", StringComparison.OrdinalIgnoreCase)).ToList());
            var usageConcrete = RecommendationFormatter.NormalizeCapabilities(
                usage.Capabilities.Where(x => !string.Equals(x, "NeedsReview", StringComparison.OrdinalIgnoreCase)).ToList());

            if (string.Equals(mode, "Rights", StringComparison.OrdinalIgnoreCase)) return rights.CloneAs("Final");
            if (string.Equals(mode, "Usage", StringComparison.OrdinalIgnoreCase)) return usage.CloneAs("Final");

            if (string.Equals(mode, "HigherOfRightsAndUsage", StringComparison.OrdinalIgnoreCase))
            {
                var merged = MergeConcreteCapabilities(rightsConcrete, usageConcrete);
                if (merged.Count > 0)
                    return RecommendationFormatter.FromCapabilities("Final", merged, rights.MatchedRuleNames.Concat(usage.MatchedRuleNames), "Combined higher-of recommendation from rights and usage.");
                if (!rights.IsNoSignal && rights.Capabilities.Count > 0) return rights.CloneAs("Final");
                return usage.CloneAs("Final");
            }

            // RightsThenUsage (default)
            if (rightsConcrete.Count > 0)
            {
                if (usageConcrete.Count > 0)
                {
                    var merged = MergeConcreteCapabilities(rightsConcrete, usageConcrete);
                    var usageAddsSomething = usage.Weight > rights.Weight || !merged.SequenceEqual(rightsConcrete, StringComparer.OrdinalIgnoreCase);
                    if (usageAddsSomething)
                        return RecommendationFormatter.FromCapabilities("Final", merged, rights.MatchedRuleNames.Concat(usage.MatchedRuleNames), "Rights-based recommendation refined with usage evidence.");
                }
                return rights.CloneAs("Final");
            }

            if (usageConcrete.Count > 0) return usage.CloneAs("Final");
            if (!rights.IsNoSignal && rights.Capabilities.Count > 0) return rights.CloneAs("Final");
            return usage.CloneAs("Final");
        }

        private static List<string> MergeConcreteCapabilities(IEnumerable<string> left, IEnumerable<string> right)
        {
            return RecommendationFormatter.NormalizeCapabilities(left.Concat(right).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
        }
    }

    public static class CoverageEvaluator
    {
        private static readonly HashSet<string> FullBaseLicenses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SalesProfessional", "SalesEnterprise", "SalesPremium",
            "CustomerServiceProfessional", "CustomerServiceEnterprise", "CustomerServicePremium",
            "FieldService"
        };

        public static CoverageResult Evaluate(SystemUserRecord user, RecommendationDecision recommendation, IReadOnlyCollection<string> actualNormalized, UserGraphLicenseSnapshot graph)
        {
            if (UserClassifier.IsSpecialAccount(user))
                return new CoverageResult { Status = "Special account", Notes = new List<string> { "This account should be reviewed separately from named end-user licensing." } };
            if (!string.Equals(graph.ActualLicenseState, "Known", StringComparison.OrdinalIgnoreCase))
                return new CoverageResult { Status = "Recommendation only", Notes = new List<string> { graph.ActualLicenseMessage } };
            if (recommendation.IsNoSignal)
                return new CoverageResult { Status = "Review", Notes = new List<string> { "No usage-based signal was available for the chosen recommendation strategy." } };
            if (recommendation.Capabilities.Count == 0 || recommendation.Capabilities.Any(c => string.Equals(c, "NeedsReview", StringComparison.OrdinalIgnoreCase)))
                return new CoverageResult { Status = "Review", Notes = new List<string> { "No confident capability recommendation. Review manually." } };
            if (actualNormalized.Count == 0)
                return new CoverageResult { Status = "Underlicensed", Notes = new List<string> { "Graph returned no assigned Dynamics SKU that matched the current rules." } };

            var missingCapabilities = recommendation.Capabilities.Where(cap => !CoversCapability(cap, actualNormalized)).ToList();
            if (missingCapabilities.Count == 0)
            {
                if (recommendation.Capabilities.Count == 1 && recommendation.Capabilities.Any(c => string.Equals(c, "TeamMembers", StringComparison.OrdinalIgnoreCase)) && HasAnyFullBase(actualNormalized))
                    return new CoverageResult { Status = "Covered", Notes = new List<string> { "Assigned license appears stronger than the current minimum recommendation." } };
                if (recommendation.Capabilities.Count > 1)
                    return new CoverageResult { Status = "Covered", Notes = new List<string> { "Multiple workloads detected. Validate base/attach manually." } };
                return new CoverageResult { Status = "Covered" };
            }

            if (actualNormalized.Any(x => x.StartsWith("Unmapped:", StringComparison.OrdinalIgnoreCase)))
                return new CoverageResult { Status = "Review", Notes = new List<string> { "One or more SKUs are not mapped by the current normalization rules." } };
            if (actualNormalized.Any(x => x.EndsWith("Attach", StringComparison.OrdinalIgnoreCase)))
                return new CoverageResult { Status = "Review", Notes = new List<string> { "Attach SKU visible but qualifying base not proven." } };

            return new CoverageResult { Status = "Underlicensed", Notes = new List<string> { "Missing: " + string.Join(", ", missingCapabilities.Select(RecommendationFormatter.CapabilityToLabel)) } };
        }

        public static bool HasAnyFullBase(IEnumerable<string> actualNormalized) => actualNormalized.Any(x => FullBaseLicenses.Contains(x));

        private static bool CoversCapability(string capability, IReadOnlyCollection<string> actual)
        {
            switch (capability)
            {
                case "TeamMembers": return actual.Any(a => string.Equals(a, "TeamMembers", StringComparison.OrdinalIgnoreCase)) || HasAnyFullBase(actual);
                case "SalesFull": return actual.Any(a => string.Equals(a, "SalesProfessional", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "SalesEnterprise", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "SalesPremium", StringComparison.OrdinalIgnoreCase)) || (actual.Any(a => string.Equals(a, "SalesAttach", StringComparison.OrdinalIgnoreCase)) && HasAnyFullBase(actual));
                case "CustomerServiceFull": return actual.Any(a => string.Equals(a, "CustomerServiceProfessional", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "CustomerServiceEnterprise", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "CustomerServicePremium", StringComparison.OrdinalIgnoreCase)) || (actual.Any(a => string.Equals(a, "CustomerServiceAttach", StringComparison.OrdinalIgnoreCase)) && HasAnyFullBase(actual));
                case "FieldServiceFull": return actual.Any(a => string.Equals(a, "FieldService", StringComparison.OrdinalIgnoreCase)) || (actual.Any(a => string.Equals(a, "FieldServiceAttach", StringComparison.OrdinalIgnoreCase)) && HasAnyFullBase(actual));
                default: return false;
            }
        }
    }

    public static class JsonHelper
    {
        public static string GetRequiredString(JsonElement element, params string[] propertyNames)
        {
            return GetString(element, propertyNames)
                   ?? throw new InvalidOperationException("Missing required string property: " + string.Join(", ", propertyNames));
        }

        public static Guid GetRequiredGuid(JsonElement element, params string[] propertyNames)
        {
            return GetGuid(element, propertyNames)
                   ?? throw new InvalidOperationException("Missing required GUID property: " + string.Join(", ", propertyNames));
        }

        public static string GetString(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null)
                    return value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
            }
            return null;
        }

        public static Guid? GetGuid(JsonElement element, params string[] propertyNames)
        {
            var text = GetString(element, propertyNames);
            return Guid.TryParse(text, out var guid) ? guid : (Guid?)null;
        }

        public static int? GetInt32(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null)
                {
                    if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number)) return number;
                    if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) return number;
                }
            }
            return null;
        }

        public static int? GetInt32FromChild(JsonElement element, string childName, string propName)
        {
            if (element.TryGetProperty(childName, out var child) && child.ValueKind == JsonValueKind.Object)
                return GetInt32(child, propName);
            return null;
        }

        public static bool? GetBoolean(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null)
                {
                    if (value.ValueKind == JsonValueKind.True) return true;
                    if (value.ValueKind == JsonValueKind.False) return false;
                    if (value.ValueKind == JsonValueKind.String && bool.TryParse(value.GetString(), out var b)) return b;
                }
            }
            return null;
        }

        public static List<JsonElement> GetArray(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.Array)
                return value.EnumerateArray().Select(x => x.Clone()).ToList();
            return new List<JsonElement>();
        }
    }

    public static class RegexHelper
    {
        public static bool IsMatch(string input, string pattern)
        {
            try { return Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant); }
            catch (ArgumentException) { return input.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0; }
        }
    }

    public static class TextHelper
    {
        public static string Shorten(string value, int maxLength = 500)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
            return normalized.Length <= maxLength ? normalized : normalized.Substring(0, maxLength);
        }
    }

    public interface IAuditLogger
    {
        void Info(string message);
        void Warn(string message);
        void Error(string message);
    }
}