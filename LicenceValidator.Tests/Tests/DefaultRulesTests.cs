using System.Linq;
using System.Reflection;
using System.IO;
using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class DefaultRulesTests
    {
        private static Ruleset _ruleset;

        [ClassInitialize]
        public static void ClassInit(TestContext ctx)
        {
            var asm = Assembly.Load("LicenceValidator");
            var resourceName = "LicenceValidator.DefaultRules.json";
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                Assert.IsNotNull(stream, "DefaultRules.json must be an EmbeddedResource in LicenceValidator.");
                using (var reader = new StreamReader(stream))
                    _ruleset = Ruleset.LoadFromJson(reader.ReadToEnd());
            }
        }

        [TestMethod]
        public void DefaultRules_IsEmbedded_CanLoad()
        {
            Assert.IsNotNull(_ruleset);
        }

        [TestMethod]
        public void DefaultRules_HasRecommendationRules()
        {
            Assert.IsTrue(_ruleset.RecommendationRules.Count > 0, "At least one recommendation rule expected.");
        }

        [TestMethod]
        public void DefaultRules_HasUsageTableProfiles()
        {
            Assert.IsTrue(_ruleset.UsageTableProfiles.Count > 0, "At least one UsageTableProfile expected.");
        }

        [TestMethod]
        public void DefaultRules_AllTableProfiles_HaveLogicalName()
        {
            foreach (var p in _ruleset.UsageTableProfiles)
                Assert.IsFalse(string.IsNullOrWhiteSpace(p.LogicalName),
                    "Every UsageTableProfile must have a LogicalName.");
        }

        [TestMethod]
        public void DefaultRules_AllRecommendationRules_HaveCapability()
        {
            foreach (var r in _ruleset.RecommendationRules)
                Assert.IsFalse(string.IsNullOrWhiteSpace(r.Capability),
                    $"Rule '{r.Name}' must have a Capability.");
        }

        [TestMethod]
        public void DefaultRules_AllRecommendationRules_HaveName()
        {
            foreach (var r in _ruleset.RecommendationRules)
                Assert.IsFalse(string.IsNullOrWhiteSpace(r.Name), "Every rule must have a Name.");
        }

        [TestMethod]
        public void DefaultRules_AllRecommendationRules_HaveAtLeastOnePattern()
        {
            foreach (var r in _ruleset.RecommendationRules)
                Assert.IsTrue(r.AnyRolePatterns != null && r.AnyRolePatterns.Count > 0,
                    $"Rule '{r.Name}' must have at least one RequiredRolePattern.");
        }

        [TestMethod]
        public void DefaultRules_NoDuplicateTableLogicalNames()
        {
            var names = _ruleset.UsageTableProfiles.Select(x => x.LogicalName.ToLowerInvariant()).ToList();
            var distinct = names.Distinct().ToList();
            Assert.AreEqual(distinct.Count, names.Count, "Duplicate LogicalNames found in UsageTableProfiles.");
        }

        [TestMethod]
        public void DefaultRules_ContainsOpportunityTable()
        {
            Assert.IsTrue(_ruleset.UsageTableProfiles.Any(x => x.LogicalName == "opportunity"),
                "Expected 'opportunity' in default UsageTableProfiles.");
        }

        [TestMethod]
        public void DefaultRules_ContainsIncidentTable()
        {
            Assert.IsTrue(_ruleset.UsageTableProfiles.Any(x => x.LogicalName == "incident"),
                "Expected 'incident' in default UsageTableProfiles.");
        }

        [TestMethod]
        public void DefaultRules_AllRecommendationRules_CapabilityRecognizedByFormatter()
        {
            var knownCapabilities = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase)
            {
                "SalesFull", "CustomerServiceFull", "FieldServiceFull", "TeamMembers"
            };
            foreach (var r in _ruleset.RecommendationRules)
                Assert.IsTrue(knownCapabilities.Contains(r.Capability),
                    $"Rule '{r.Name}' has Capability '{r.Capability}' which is not recognized by RecommendationFormatter. Use one of: {string.Join(", ", knownCapabilities)}.");
        }

        [TestMethod]
        public void DefaultRules_AllUsageTableProfiles_CapabilityRecognizedByFormatter()
        {
            var knownCapabilities = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase)
            {
                "SalesFull", "CustomerServiceFull", "FieldServiceFull", "TeamMembers"
            };
            foreach (var p in _ruleset.UsageTableProfiles)
                Assert.IsTrue(knownCapabilities.Contains(p.Capability),
                    $"UsageTableProfile '{p.LogicalName}' has Capability '{p.Capability}' which is not recognized by RecommendationFormatter. Use one of: {string.Join(", ", knownCapabilities)}.");
        }

        [TestMethod]
        public void DefaultRules_OmnichannelRole_MatchesCustomerServiceRule()
        {
            var evidence = new UserEvidence { RoleNames = new System.Collections.Generic.List<string> { "CRM_Omnichannel_Supervisor" } };
            var decision = _ruleset.EvaluateRights(evidence);
            Assert.IsFalse(decision.IsReviewOnly, "Omnichannel role should match a Customer Service rule.");
            CollectionAssert.Contains(decision.Capabilities, "CustomerServiceFull");
        }

        [TestMethod]
        public void DefaultRules_SalespersonRole_MatchesSalesRule()
        {
            var evidence = new UserEvidence { RoleNames = new System.Collections.Generic.List<string> { "Salesperson" } };
            var decision = _ruleset.EvaluateRights(evidence);
            Assert.IsFalse(decision.IsReviewOnly);
            CollectionAssert.Contains(decision.Capabilities, "SalesFull");
        }

        [TestMethod]
        public void DefaultRules_FieldServiceRole_MatchesFieldServiceRule()
        {
            var evidence = new UserEvidence { RoleNames = new System.Collections.Generic.List<string> { "Field Service - Dispatcher" } };
            var decision = _ruleset.EvaluateRights(evidence);
            Assert.IsFalse(decision.IsReviewOnly);
            CollectionAssert.Contains(decision.Capabilities, "FieldServiceFull");
        }

        [TestMethod]
        public void DefaultRules_SalesRule_RecommendedSkuIsNotNeedsReview()
        {
            var evidence = new UserEvidence { RoleNames = new System.Collections.Generic.List<string> { "Salesperson" } };
            var decision = _ruleset.EvaluateRights(evidence);
            Assert.AreNotEqual("Needs review", decision.RecommendedSku,
                "RecommendedSku should resolve to a known SKU string, not 'Needs review'.");
        }
    }
}

