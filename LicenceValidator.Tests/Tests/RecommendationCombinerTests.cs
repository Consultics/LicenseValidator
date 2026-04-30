using System.Collections.Generic;
using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class RecommendationCombinerTests
    {
        private static RecommendationDecision Sales =>
            RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesEnterprise" });
        private static RecommendationDecision SalesFull =>
            RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesFull" });
        private static RecommendationDecision CS =>
            RecommendationFormatter.FromCapabilities("Usage", new[] { "CustomerServiceFull" });
        private static RecommendationDecision TeamMembers =>
            RecommendationFormatter.FromCapabilities("Rights", new[] { "TeamMembers" });
        private static RecommendationDecision Review =>
            RecommendationDecision.CreateReview("Rights", "No match");
        private static RecommendationDecision NoSignal =>
            RecommendationDecision.CreateNoSignal("Usage", "No signal");

        // ── Mode: Rights ──────────────────────────────────────────────────────
        [TestMethod]
        public void Combine_ModeRights_ReturnsRights()
        {
            var result = RecommendationCombiner.Combine(Sales, CS, RecommendationModes.Rights);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "SalesEnterprise");
        }

        // ── Mode: Usage ───────────────────────────────────────────────────────
        [TestMethod]
        public void Combine_ModeUsage_ReturnsUsage()
        {
            var result = RecommendationCombiner.Combine(Sales, CS, RecommendationModes.Usage);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "CustomerServiceFull");
        }

        // ── Mode: RightsThenUsage ─────────────────────────────────────────────
        [TestMethod]
        public void Combine_RightsThenUsage_RightsConcreteOnly_ReturnsRights()
        {
            var result = RecommendationCombiner.Combine(Sales, NoSignal, RecommendationModes.RightsThenUsage);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "SalesEnterprise");
        }

        [TestMethod]
        public void Combine_RightsThenUsage_BothConcrete_MergesSame_ReturnsRights()
        {
            var result = RecommendationCombiner.Combine(Sales, Sales, RecommendationModes.RightsThenUsage);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "SalesEnterprise");
        }

        [TestMethod]
        public void Combine_RightsThenUsage_BothConcreteDifferent_MergesBoth()
        {
            var result = RecommendationCombiner.Combine(SalesFull, CS, RecommendationModes.RightsThenUsage);
            Assert.AreEqual("Final", result.Source);
            Assert.IsTrue(result.Capabilities.Count >= 1);
        }

        [TestMethod]
        public void Combine_RightsThenUsage_OnlyUsageConcrete_ReturnsUsage()
        {
            var result = RecommendationCombiner.Combine(Review, CS, RecommendationModes.RightsThenUsage);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "CustomerServiceFull");
        }

        [TestMethod]
        public void Combine_RightsThenUsage_BothReview_ReturnsFinal()
        {
            var result = RecommendationCombiner.Combine(Review, Review, RecommendationModes.RightsThenUsage);
            Assert.AreEqual("Final", result.Source);
        }

        // ── Mode: HigherOfRightsAndUsage ──────────────────────────────────────
        [TestMethod]
        public void Combine_HigherOf_BothConcrete_ReturnsMerged()
        {
            var result = RecommendationCombiner.Combine(SalesFull, CS, RecommendationModes.HigherOfRightsAndUsage);
            Assert.AreEqual("Final", result.Source);
            Assert.IsTrue(result.Capabilities.Count >= 1);
        }

        [TestMethod]
        public void Combine_HigherOf_OnlyRightsConcrete_ReturnsRights()
        {
            var result = RecommendationCombiner.Combine(Sales, NoSignal, RecommendationModes.HigherOfRightsAndUsage);
            Assert.AreEqual("Final", result.Source);
            CollectionAssert.Contains(result.Capabilities, "SalesEnterprise");
        }

        // ── CloneAs ───────────────────────────────────────────────────────────
        [TestMethod]
        public void CloneAs_ChangesSource_PreservesCapabilities()
        {
            var original = RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesEnterprise" });
            var clone = original.CloneAs("Final");
            Assert.AreEqual("Final", clone.Source);
            CollectionAssert.Contains(clone.Capabilities, "SalesEnterprise");
        }

        [TestMethod]
        public void CloneAs_DoesNotMutateOriginal()
        {
            var original = RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesEnterprise" });
            var clone = original.CloneAs("Final");
            clone.Capabilities.Add("TeamMembers");
            Assert.AreEqual(1, original.Capabilities.Count);
        }
    }
}
