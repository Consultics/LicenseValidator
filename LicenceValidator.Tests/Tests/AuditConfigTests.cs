using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class AuditConfigTests
    {
        // ── UsageEnabledEffective ─────────────────────────────────────────────

        [TestMethod]
        public void UsageEnabledEffective_RightsOnly_ReturnsFalse()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.RightsOnly, UsageEnabled = true };
            Assert.IsFalse(cfg.UsageEnabledEffective);
        }

        [TestMethod]
        public void UsageEnabledEffective_RightsAndUsage_UsageFalse_ReturnsFalse()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.RightsAndUsage, UsageEnabled = false };
            Assert.IsFalse(cfg.UsageEnabledEffective);
        }

        [TestMethod]
        public void UsageEnabledEffective_RightsAndUsage_UsageTrue_ReturnsTrue()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.RightsAndUsage, UsageEnabled = true };
            Assert.IsTrue(cfg.UsageEnabledEffective);
        }

        [TestMethod]
        public void UsageEnabledEffective_NoGraphRightsAndUsage_UsageTrue_ReturnsTrue()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.NoGraphRightsAndUsage, UsageEnabled = true };
            Assert.IsTrue(cfg.UsageEnabledEffective);
        }

        [TestMethod]
        public void UsageEnabledEffective_NoGraphRightsOnly_ReturnsFalse()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.NoGraphRightsOnly, UsageEnabled = true };
            Assert.IsFalse(cfg.UsageEnabledEffective);
        }

        // ── GraphEnabled ──────────────────────────────────────────────────────

        [TestMethod]
        public void GraphEnabled_RightsAndUsage_ReturnsTrue()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.RightsAndUsage };
            Assert.IsTrue(cfg.GraphEnabled);
        }

        [TestMethod]
        public void GraphEnabled_NoGraphRightsOnly_ReturnsFalse()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.NoGraphRightsOnly };
            Assert.IsFalse(cfg.GraphEnabled);
        }

        [TestMethod]
        public void GraphEnabled_NoGraphRightsAndUsage_ReturnsFalse()
        {
            var cfg = new ToolConfig { AuditMode = AppModes.NoGraphRightsAndUsage };
            Assert.IsFalse(cfg.GraphEnabled);
        }

        // ── AppModes static helpers ───────────────────────────────────────────

        [TestMethod]
        public void AppModes_UsesUsage_RightsOnly_ReturnsFalse()
        {
            Assert.IsFalse(AppModes.UsesUsage(AppModes.RightsOnly));
        }

        [TestMethod]
        public void AppModes_UsesUsage_RightsAndUsage_ReturnsTrue()
        {
            Assert.IsTrue(AppModes.UsesUsage(AppModes.RightsAndUsage));
        }

        [TestMethod]
        public void AppModes_UsesUsage_NoGraphRightsAndUsage_ReturnsTrue()
        {
            Assert.IsTrue(AppModes.UsesUsage(AppModes.NoGraphRightsAndUsage));
        }

        [TestMethod]
        public void AppModes_UsesGraph_RightsAndUsage_ReturnsTrue()
        {
            Assert.IsTrue(AppModes.UsesGraph(AppModes.RightsAndUsage));
        }

        [TestMethod]
        public void AppModes_UsesGraph_NoGraphRightsOnly_ReturnsFalse()
        {
            Assert.IsFalse(AppModes.UsesGraph(AppModes.NoGraphRightsOnly));
        }

        // ── EffectiveRecommendationMode ───────────────────────────────────────

        [TestMethod]
        public void EffectiveRecommendationMode_WithGraph_UsesGraphMode()
        {
            var cfg = new ToolConfig
            {
                AuditMode = AppModes.RightsAndUsage,
                RecommendationModeWithGraph = RecommendationModes.RightsThenUsage,
                RecommendationModeWithoutGraph = RecommendationModes.Rights
            };
            Assert.AreEqual(RecommendationModes.RightsThenUsage, cfg.EffectiveRecommendationMode);
        }

        [TestMethod]
        public void EffectiveRecommendationMode_WithoutGraph_UsesNoGraphMode()
        {
            var cfg = new ToolConfig
            {
                AuditMode = AppModes.NoGraphRightsOnly,
                RecommendationModeWithGraph = RecommendationModes.RightsThenUsage,
                RecommendationModeWithoutGraph = RecommendationModes.Rights
            };
            Assert.AreEqual(RecommendationModes.Rights, cfg.EffectiveRecommendationMode);
        }

        // ── Defaults ─────────────────────────────────────────────────────────

        [TestMethod]
        public void ToolConfig_Defaults_AuditModeIsRightsAndUsage()
        {
            var cfg = new ToolConfig();
            Assert.AreEqual(AppModes.RightsAndUsage, cfg.AuditMode);
        }

        [TestMethod]
        public void ToolConfig_Defaults_UsageEnabledIsTrue()
        {
            var cfg = new ToolConfig();
            Assert.IsTrue(cfg.UsageEnabled);
        }
    }
}
