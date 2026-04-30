using System;
using System.Collections.Generic;
using System.Text.Json;
using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class UserClassifierTests
    {
        // ── DetermineUserType ─────────────────────────────────────────────────
        [TestMethod]
        public void DetermineUserType_ApplicationId_ReturnsApplicationUser()
        {
            var u = new SystemUserRecord { ApplicationId = Guid.NewGuid() };
            Assert.AreEqual("ApplicationUser", UserClassifier.DetermineUserType(u));
        }

        [TestMethod]
        public void DetermineUserType_AccessMode0_ReturnsHuman()
        {
            var u = new SystemUserRecord { AccessMode = 0 };
            Assert.AreEqual("Human", UserClassifier.DetermineUserType(u));
        }

        [TestMethod]
        public void DetermineUserType_AccessMode1_ReturnsAdministrative()
        {
            var u = new SystemUserRecord { AccessMode = 1 };
            Assert.AreEqual("Administrative", UserClassifier.DetermineUserType(u));
        }

        [TestMethod]
        public void DetermineUserType_AccessMode4_ReturnsNonInteractive()
        {
            var u = new SystemUserRecord { AccessMode = 4 };
            Assert.AreEqual("NonInteractive", UserClassifier.DetermineUserType(u));
        }

        [TestMethod]
        public void DetermineUserType_AccessMode2_ReturnsReadOnly()
        {
            var u = new SystemUserRecord { AccessMode = 2 };
            Assert.AreEqual("ReadOnly", UserClassifier.DetermineUserType(u));
        }

        // ── IsApplicationUser ─────────────────────────────────────────────────
        [TestMethod]
        public void IsApplicationUser_WithApplicationId_ReturnsTrue()
        {
            var u = new SystemUserRecord { ApplicationId = Guid.NewGuid() };
            Assert.IsTrue(UserClassifier.IsApplicationUser(u));
        }

        [TestMethod]
        public void IsApplicationUser_UserTypeString_ReturnsTrue()
        {
            var u = new SystemUserRecord { UserType = "ApplicationUser" };
            Assert.IsTrue(UserClassifier.IsApplicationUser(u));
        }

        [TestMethod]
        public void IsApplicationUser_HumanUser_ReturnsFalse()
        {
            var u = new SystemUserRecord { UserType = "Human" };
            Assert.IsFalse(UserClassifier.IsApplicationUser(u));
        }

        // ── IsSpecialAccount ──────────────────────────────────────────────────
        [TestMethod]
        public void IsSpecialAccount_NonInteractive_ReturnsTrue()
        {
            var u = new SystemUserRecord { UserType = "NonInteractive" };
            Assert.IsTrue(UserClassifier.IsSpecialAccount(u));
        }

        [TestMethod]
        public void IsSpecialAccount_Human_ReturnsFalse()
        {
            var u = new SystemUserRecord { UserType = "Human" };
            Assert.IsFalse(UserClassifier.IsSpecialAccount(u));
        }

        [TestMethod]
        public void IsSpecialAccount_DelegatedAdmin_ReturnsTrue()
        {
            var u = new SystemUserRecord { UserType = "DelegatedAdmin" };
            Assert.IsTrue(UserClassifier.IsSpecialAccount(u));
        }
    }

    [TestClass]
    public class RecommendationFormatterTests
    {
        // ── CapabilityRank ────────────────────────────────────────────────────
        [TestMethod]
        public void CapabilityRank_TeamMembers_IsLowerThanSalesFull()
        {
            Assert.IsTrue(RecommendationFormatter.CapabilityRank("SalesFull") > RecommendationFormatter.CapabilityRank("TeamMembers"));
        }

        [TestMethod]
        public void CapabilityRank_Unknown_ReturnsZero()
        {
            Assert.AreEqual(0, RecommendationFormatter.CapabilityRank("Unknown"));
        }

        // ── CapabilityToLabel ─────────────────────────────────────────────────
        [TestMethod]
        public void CapabilityToLabel_TeamMembers_ReturnsReadableLabel()
        {
            Assert.AreEqual("Team Members", RecommendationFormatter.CapabilityToLabel("TeamMembers"));
        }

        [TestMethod]
        public void CapabilityToLabel_SalesFull_ReturnsReadableLabel()
        {
            Assert.AreEqual("Sales", RecommendationFormatter.CapabilityToLabel("SalesFull"));
        }

        // ── ToPreferredSku ────────────────────────────────────────────────────
        [TestMethod]
        public void ToPreferredSku_TeamMembers_ReturnsSku()
        {
            Assert.AreEqual("TeamMembers", RecommendationFormatter.ToPreferredSku(new[] { "TeamMembers" }));
        }

        [TestMethod]
        public void ToPreferredSku_EmptyList_ReturnsNeedsReview()
        {
            Assert.AreEqual("Needs review", RecommendationFormatter.ToPreferredSku(new string[0]));
        }

        // ── FromCapabilities ──────────────────────────────────────────────────
        [TestMethod]
        public void FromCapabilities_SingleCapability_IsNotReviewOnly()
        {
            var d = RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesEnterprise" });
            Assert.IsFalse(d.IsReviewOnly);
            Assert.IsFalse(d.IsNoSignal);
        }

        [TestMethod]
        public void FromCapabilities_EmptyList_IsReviewOnly()
        {
            var d = RecommendationFormatter.FromCapabilities("Rights", new string[0]);
            Assert.IsTrue(d.IsReviewOnly);
        }

        [TestMethod]
        public void FromCapabilities_WeightIsPositive_ForKnownCapability()
        {
            var d = RecommendationFormatter.FromCapabilities("Rights", new[] { "SalesFull" });
            Assert.IsTrue(d.Weight > 0);
        }
    }

    [TestClass]
    public class RecommendationDecisionTests
    {
        [TestMethod]
        public void CreateNoSignal_HasIsNoSignalTrue()
        {
            var d = RecommendationDecision.CreateNoSignal("Usage", "No data");
            Assert.IsTrue(d.IsNoSignal);
            Assert.IsFalse(d.IsReviewOnly);
        }

        [TestMethod]
        public void CreateReview_HasIsReviewOnlyTrue()
        {
            var d = RecommendationDecision.CreateReview("Rights", "No match");
            Assert.IsTrue(d.IsReviewOnly);
            Assert.IsFalse(d.IsNoSignal);
        }

        [TestMethod]
        public void CreateReview_WithRuleNames_PopulatesMatchedRuleNames()
        {
            var d = RecommendationDecision.CreateReview("Rights", "msg", new[] { "Rule1", "Rule2" });
            Assert.AreEqual(2, d.MatchedRuleNames.Count);
        }

        [TestMethod]
        public void CreateReview_NullRuleNames_MatchedRuleNamesEmpty()
        {
            var d = RecommendationDecision.CreateReview("Rights", "msg", null);
            Assert.AreEqual(0, d.MatchedRuleNames.Count);
        }
    }

    [TestClass]
    public class JsonHelperTests
    {
        private static JsonElement Parse(string json) =>
            JsonDocument.Parse(json).RootElement;

        [TestMethod]
        public void GetString_ExistingProperty_ReturnsValue()
        {
            var el = Parse("{\"name\":\"Alice\"}");
            Assert.AreEqual("Alice", JsonHelper.GetString(el, "name"));
        }

        [TestMethod]
        public void GetString_MissingProperty_ReturnsNull()
        {
            var el = Parse("{\"other\":\"x\"}");
            Assert.IsNull(JsonHelper.GetString(el, "name"));
        }

        [TestMethod]
        public void GetString_MultipleNames_ReturnsFirstMatch()
        {
            var el = Parse("{\"b\":\"second\"}");
            Assert.AreEqual("second", JsonHelper.GetString(el, "a", "b"));
        }

        [TestMethod]
        public void GetGuid_ValidGuid_ReturnsGuid()
        {
            var id = Guid.NewGuid();
            var el = Parse("{\"id\":\"" + id.ToString() + "\"}");
            Assert.AreEqual(id, JsonHelper.GetGuid(el, "id"));
        }

        [TestMethod]
        public void GetGuid_InvalidGuid_ReturnsNull()
        {
            var el = Parse("{\"id\":\"not-a-guid\"}");
            Assert.IsNull(JsonHelper.GetGuid(el, "id"));
        }

        [TestMethod]
        public void GetInt32_IntegerProperty_ReturnsValue()
        {
            var el = Parse("{\"count\":42}");
            Assert.AreEqual(42, JsonHelper.GetInt32(el, "count"));
        }

        [TestMethod]
        public void GetInt32_StringNumber_ReturnsValue()
        {
            var el = Parse("{\"count\":\"99\"}");
            Assert.AreEqual(99, JsonHelper.GetInt32(el, "count"));
        }

        [TestMethod]
        public void GetInt32_MissingProperty_ReturnsNull()
        {
            var el = Parse("{\"other\":1}");
            Assert.IsNull(JsonHelper.GetInt32(el, "count"));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void GetRequiredString_MissingProperty_Throws()
        {
            var el = Parse("{\"other\":\"x\"}");
            JsonHelper.GetRequiredString(el, "name");
        }

        [TestMethod]
        public void GetRequiredString_ExistingProperty_ReturnsValue()
        {
            var el = Parse("{\"name\":\"Bob\"}");
            Assert.AreEqual("Bob", JsonHelper.GetRequiredString(el, "name"));
        }

        [TestMethod]
        public void GetBoolean_TrueValue_ReturnsTrue()
        {
            var el = Parse("{\"active\":true}");
            Assert.IsTrue(JsonHelper.GetBoolean(el, "active"));
        }

        [TestMethod]
        public void GetBoolean_FalseValue_ReturnsFalse()
        {
            var el = Parse("{\"active\":false}");
            Assert.IsFalse(JsonHelper.GetBoolean(el, "active"));
        }

        [TestMethod]
        public void GetBoolean_StringTrue_ReturnsTrue()
        {
            var el = Parse("{\"active\":\"true\"}");
            Assert.IsTrue(JsonHelper.GetBoolean(el, "active"));
        }

        [TestMethod]
        public void GetBoolean_MissingProperty_ReturnsNull()
        {
            var el = Parse("{\"other\":true}");
            Assert.IsNull(JsonHelper.GetBoolean(el, "active"));
        }
    }
}
