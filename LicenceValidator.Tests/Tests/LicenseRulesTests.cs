using System.Collections.Generic;
using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class LicenseRulesTests
    {
        // ── LoadFromJson ──────────────────────────────────────────────────────

        [TestMethod]
        public void LoadFromJson_ValidEmptyObject_ReturnsRuleset()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            Assert.IsNotNull(ruleset);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void LoadFromJson_EmptyString_Throws()
        {
            Ruleset.LoadFromJson("");
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void LoadFromJson_WhitespaceOnly_Throws()
        {
            Ruleset.LoadFromJson("   ");
        }

        [TestMethod]
        public void LoadFromJson_MissingCollections_InitializesEmpty()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            Assert.IsNotNull(ruleset.RecommendationRules);
            Assert.IsNotNull(ruleset.UsageTableProfiles);
            Assert.IsNotNull(ruleset.LicenseNormalization);
            Assert.IsNotNull(ruleset.UsageExcludeEntityPatterns);
            Assert.IsNotNull(ruleset.UsageIncludeEntityPatterns);
        }

        // ── NormalizeAssignedSkus ─────────────────────────────────────────────

        [TestMethod]
        public void NormalizeAssignedSkus_D365Sales_ReturnsSalesEnterprise()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "DYN365_ENTERPRISE_SALES" });
            CollectionAssert.Contains(result, "SalesEnterprise");
        }

        [TestMethod]
        public void NormalizeAssignedSkus_D365SalesAlias_ReturnsSalesEnterprise()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "D365_SALES_ENT" });
            CollectionAssert.Contains(result, "SalesEnterprise");
        }

        [TestMethod]
        public void NormalizeAssignedSkus_TeamMembers_ReturnsMapped()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "DYN365_ENTERPRISE_TEAM_MEMBERS" });
            CollectionAssert.Contains(result, "TeamMembers");
        }

        [TestMethod]
        public void NormalizeAssignedSkus_Office365E3_IsFiltered()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "ENTERPRISEPACK" });
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_ExchangeStandard_IsFiltered()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "EXCHANGESTANDARD" });
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_PowerAppsDev_IsFiltered()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "POWERAPPS_DEV" });
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_CcibotsViral_IsFiltered()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "CCIBOTS_PRIVPREV_VIRAL" });
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_UnknownSku_ReturnsUnmappedPrefix()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new[] { "TOTALLY_UNKNOWN_SKU_XYZ" });
            Assert.AreEqual(1, result.Count);
            StringAssert.StartsWith(result[0], "Unmapped:");
        }

        [TestMethod]
        public void NormalizeAssignedSkus_EmptyInput_ReturnsEmpty()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var result = ruleset.NormalizeAssignedSkus(new string[0]);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_TwoAliasesForSameSku_Deduplicated()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            // DYN365_ENTERPRISE_SALES and D365_SALES_ENT both map to SalesEnterprise
            var result = ruleset.NormalizeAssignedSkus(new[] { "DYN365_ENTERPRISE_SALES", "D365_SALES_ENT" });
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("SalesEnterprise", result[0]);
        }

        [TestMethod]
        public void NormalizeAssignedSkus_CustomRule_OverridesBuiltIn()
        {
            var json = "{\"LicenseNormalization\":[{\"Pattern\":\"DYN365_ENTERPRISE_SALES\",\"Normalized\":\"CustomSales\"}]}";
            var ruleset = Ruleset.LoadFromJson(json);
            var result = ruleset.NormalizeAssignedSkus(new[] { "DYN365_ENTERPRISE_SALES" });
            CollectionAssert.Contains(result, "CustomSales");
        }

        // ── BuiltInNormalization – no duplicate keys ───────────────────────────

        [TestMethod]
        public void BuiltInNormalization_NoDuplicateKeys_DoesNotThrowOnInit()
        {
            // Duplicate keys in the static Dictionary initializer throw ArgumentException.
            // Successfully loading a Ruleset proves the static init succeeded.
            var ruleset = Ruleset.LoadFromJson("{}");
            Assert.IsNotNull(ruleset);
        }

        // ── EvaluateRights ────────────────────────────────────────────────────

        [TestMethod]
        public void EvaluateRights_NoRules_IsReviewOnly()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var evidence = new UserEvidence { RoleNames = new List<string> { "Salesperson" } };
            var decision = ruleset.EvaluateRights(evidence);
            Assert.IsTrue(decision.IsReviewOnly);
        }

        [TestMethod]
        public void EvaluateRights_MatchingRolePattern_ReturnsCapability()
        {
            var json = "{\"RecommendationRules\":[{\"Name\":\"Sales\",\"Priority\":100,\"Capability\":\"SalesEnterprise\",\"AnyRolePatterns\":[\"(?i)salesperson\"]}]}";
            var ruleset = Ruleset.LoadFromJson(json);
            var evidence = new UserEvidence { RoleNames = new List<string> { "Salesperson" } };
            var decision = ruleset.EvaluateRights(evidence);
            Assert.IsFalse(decision.IsReviewOnly);
            CollectionAssert.Contains(decision.Capabilities, "SalesEnterprise");
        }

        [TestMethod]
        public void EvaluateRights_NonMatchingRole_IsReviewOnly()
        {
            var json = "{\"RecommendationRules\":[{\"Name\":\"Sales\",\"Priority\":100,\"Capability\":\"SalesEnterprise\",\"AnyRolePatterns\":[\"(?i)salesperson\"]}]}";
            var ruleset = Ruleset.LoadFromJson(json);
            var evidence = new UserEvidence { RoleNames = new List<string> { "System Administrator" } };
            var decision = ruleset.EvaluateRights(evidence);
            Assert.IsTrue(decision.IsReviewOnly);
        }

        [TestMethod]
        public void EvaluateRights_EmptyEvidence_IsReviewOnly()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var decision = ruleset.EvaluateRights(new UserEvidence());
            Assert.IsTrue(decision.IsReviewOnly);
        }

        // ── EvaluateUsage ─────────────────────────────────────────────────────

        [TestMethod]
        public void EvaluateUsage_NoSignal_IsNoSignal()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var usage = new UserUsageSummary(); // HasAnyBusinessDataSignal == false
            var decision = ruleset.EvaluateUsage(usage);
            Assert.IsTrue(decision.IsNoSignal);
        }

        [TestMethod]
        public void EvaluateUsage_WithSignalNoCapability_IsReviewOnly()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var usage = new UserUsageSummary
            {
                BusinessSignalRecordCount = 10
                // DetectedCapabilities is empty
            };
            var decision = ruleset.EvaluateUsage(usage);
            Assert.IsTrue(decision.IsReviewOnly);
            Assert.IsFalse(decision.IsNoSignal);
        }

        [TestMethod]
        public void EvaluateUsage_WithCapability_ReturnsCapabilityInDecision()
        {
            var ruleset = Ruleset.LoadFromJson("{}");
            var usage = new UserUsageSummary
            {
                BusinessSignalRecordCount = 5,
                DetectedCapabilities = new List<string> { "SalesEnterprise" }
            };
            var decision = ruleset.EvaluateUsage(usage);
            Assert.IsFalse(decision.IsReviewOnly);
            Assert.IsFalse(decision.IsNoSignal);
            CollectionAssert.Contains(decision.Capabilities, "SalesEnterprise");
        }
    }
}

