using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class SettingsTests
    {
        // ── Settings ──────────────────────────────────────────────────────────

        [TestMethod]
        public void Settings_Defaults_ActiveProfileIsDefault()
        {
            var s = new Settings();
            Assert.AreEqual("Default", s.ActiveProfileName);
        }

        [TestMethod]
        public void Settings_Defaults_ProfilesListNotNull()
        {
            var s = new Settings();
            Assert.IsNotNull(s.Profiles);
        }

        [TestMethod]
        public void GetActiveProfile_NoProfiles_CreatesDefault()
        {
            var s = new Settings();
            var profile = s.GetActiveProfile();
            Assert.IsNotNull(profile);
            Assert.AreEqual("Default", profile.Name);
        }

        [TestMethod]
        public void GetActiveProfile_ExistingProfile_ReturnsSame()
        {
            var s = new Settings { ActiveProfileName = "Prod" };
            s.Profiles.Add(new SettingsProfile { Name = "Prod", TenantId = "tid1" });
            var profile = s.GetActiveProfile();
            Assert.AreEqual("tid1", profile.TenantId);
        }

        [TestMethod]
        public void GetActiveProfile_CaseInsensitive_ReturnsMatch()
        {
            var s = new Settings { ActiveProfileName = "PROD" };
            s.Profiles.Add(new SettingsProfile { Name = "prod", TenantId = "tid-case" });
            var profile = s.GetActiveProfile();
            Assert.AreEqual("tid-case", profile.TenantId);
        }

        [TestMethod]
        public void GetActiveProfile_MissingProfile_CreatesAndAdds()
        {
            var s = new Settings { ActiveProfileName = "Missing" };
            var profile = s.GetActiveProfile();
            Assert.AreEqual("Missing", profile.Name);
            Assert.AreEqual(1, s.Profiles.Count);
        }

        // ── SettingsProfile ───────────────────────────────────────────────────

        [TestMethod]
        public void SettingsProfile_Defaults_AuditModeIsRightsAndUsage()
        {
            var p = new SettingsProfile();
            Assert.AreEqual("RightsAndUsage", p.AuditMode);
        }

        [TestMethod]
        public void SettingsProfile_Defaults_UsageEnabledIsTrue()
        {
            var p = new SettingsProfile();
            Assert.IsTrue(p.UsageEnabled);
        }

        [TestMethod]
        public void SettingsProfile_Defaults_ActivityLookbackIs180()
        {
            var p = new SettingsProfile();
            Assert.AreEqual(180, p.ActivityLookbackDays);
        }

        [TestMethod]
        public void SettingsProfile_Defaults_RecommendationModesAreRightsThenUsage()
        {
            var p = new SettingsProfile();
            Assert.AreEqual("RightsThenUsage", p.RecommendationModeWithGraph);
            Assert.AreEqual("RightsThenUsage", p.RecommendationModeWithoutGraph);
        }

        [TestMethod]
        public void SettingsProfile_Defaults_GraphOptionalIsTrue()
        {
            var p = new SettingsProfile();
            Assert.IsTrue(p.GraphOptional);
        }
    }
}
