using System.Collections.Generic;
using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class CoverageEvaluatorTests
    {
        private static SystemUserRecord HumanUser() => new SystemUserRecord { UserType = "Human", AccessMode = 0 };
        private static UserGraphLicenseSnapshot KnownGraph() => new UserGraphLicenseSnapshot { ActualLicenseState = "Known", ActualLicenseMessage = "ok", EnabledByMode = true };
        private static UserGraphLicenseSnapshot UnknownGraph() => UserGraphLicenseSnapshot.CreateUnknown("Skipped", "Graph nicht aktiviert", false);
        private static RecommendationDecision SalesFullDecision() => RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesFull" });
        private static RecommendationDecision CSFullDecision() => RecommendationFormatter.FromCapabilities("Rights", new[] { "CustomerServiceFull" });
        private static RecommendationDecision TeamMembersDecision() => RecommendationFormatter.FromCapabilities("Rights", new[] { "TeamMembers" });
        private static RecommendationDecision FieldServiceDecision() => RecommendationFormatter.FromCapabilities("Rights", new[] { "FieldServiceFull" });
        private static RecommendationDecision ReviewDecision() => RecommendationDecision.CreateReview("Rights", "No match");
        private static RecommendationDecision NoSignalDecision() => RecommendationDecision.CreateNoSignal("Usage", "No signal");

        // ── Special account ───────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_ApplicationUser_ReturnsSpecialAccount()
        {
            var user = new SystemUserRecord { ApplicationId = System.Guid.NewGuid() };
            var result = CoverageEvaluator.Evaluate(user, SalesFullDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Special account", result.Status);
        }

        [TestMethod]
        public void Evaluate_NonInteractiveUser_ReturnsSpecialAccount()
        {
            var user = new SystemUserRecord { UserType = "NonInteractive", AccessMode = 4 };
            var result = CoverageEvaluator.Evaluate(user, SalesFullDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Special account", result.Status);
        }

        // ── Graph unknown ─────────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_GraphUnknown_ReturnsRecommendationOnly()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "SalesEnterprise" }, UnknownGraph());
            Assert.AreEqual("Recommendation only", result.Status);
        }

        // ── No signal ─────────────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_NoSignalDecision_ReturnsReview()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), NoSignalDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Review", result.Status);
        }

        // ── Review decision ───────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_ReviewDecision_ReturnsReview()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), ReviewDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Review", result.Status);
        }

        // ── No actual licenses ────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_NoActualLicenses_ReturnsUnderlicensed()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new string[0], KnownGraph());
            Assert.AreEqual("Underlicensed", result.Status);
        }

        // ── Covered ───────────────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_SalesFullWithSalesEnterprise_ReturnsCovered()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
        }

        [TestMethod]
        public void Evaluate_SalesFullWithSalesProfessional_ReturnsCovered()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "SalesProfessional" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
        }

        [TestMethod]
        public void Evaluate_TeamMembersWithFullBase_ReturnsCoveredWithNote()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), TeamMembersDecision(), new[] { "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
            Assert.IsTrue(result.Notes.Count > 0);
        }

        [TestMethod]
        public void Evaluate_TeamMembersWithTeamMembersLicense_ReturnsCovered()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), TeamMembersDecision(), new[] { "TeamMembers" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
        }

        [TestMethod]
        public void Evaluate_FieldServiceViaAttachAndBase_ReturnsCovered()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), FieldServiceDecision(), new[] { "FieldServiceAttach", "SalesEnterprise" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
        }

        [TestMethod]
        public void Evaluate_CSFullWithCSEnterprise_ReturnsCovered()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), CSFullDecision(), new[] { "CustomerServiceEnterprise" }, KnownGraph());
            Assert.AreEqual("Covered", result.Status);
        }

        // ── Underlicensed ─────────────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_SalesFullWithTeamMembersOnly_ReturnsUnderlicensed()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "TeamMembers" }, KnownGraph());
            Assert.AreEqual("Underlicensed", result.Status);
        }

        // ── Unmapped SKU → Review ─────────────────────────────────────────────
        [TestMethod]
        public void Evaluate_UnmappedSku_ReturnsReview()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "Unmapped:SOMESKU" }, KnownGraph());
            Assert.AreEqual("Review", result.Status);
        }

        // ── Attach without base → Review ──────────────────────────────────────
        [TestMethod]
        public void Evaluate_AttachWithoutBase_ReturnsReview()
        {
            var result = CoverageEvaluator.Evaluate(HumanUser(), SalesFullDecision(), new[] { "SalesAttach" }, KnownGraph());
            Assert.AreEqual("Review", result.Status);
        }

        // ── HasAnyFullBase ────────────────────────────────────────────────────
        [TestMethod]
        public void HasAnyFullBase_WithSalesEnterprise_ReturnsTrue()
        {
            Assert.IsTrue(CoverageEvaluator.HasAnyFullBase(new[] { "SalesEnterprise" }));
        }

        [TestMethod]
        public void HasAnyFullBase_WithTeamMembers_ReturnsFalse()
        {
            Assert.IsFalse(CoverageEvaluator.HasAnyFullBase(new[] { "TeamMembers" }));
        }

        [TestMethod]
        public void HasAnyFullBase_WithFieldService_ReturnsTrue()
        {
            Assert.IsTrue(CoverageEvaluator.HasAnyFullBase(new[] { "FieldService" }));
        }
    }
}
