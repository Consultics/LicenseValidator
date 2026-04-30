using LicenceValidator.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LicenceValidator.Tests
{
    [TestClass]
    public class AuditHelpersTests
    {
        // ── RegexHelper ───────────────────────────────────────────────────────

        [TestMethod]
        public void RegexHelper_IsMatch_ValidPattern_ReturnsTrue()
        {
            Assert.IsTrue(RegexHelper.IsMatch("Salesperson", "(?i)sales"));
        }

        [TestMethod]
        public void RegexHelper_IsMatch_CaseInsensitive_ReturnsTrue()
        {
            Assert.IsTrue(RegexHelper.IsMatch("SALESPERSON", "salesperson"));
        }

        [TestMethod]
        public void RegexHelper_IsMatch_NoMatch_ReturnsFalse()
        {
            Assert.IsFalse(RegexHelper.IsMatch("System Administrator", "(?i)salesperson"));
        }

        [TestMethod]
        public void RegexHelper_IsMatch_InvalidRegex_FallsBackToContains()
        {
            // "[invalid" is invalid regex; should fall back to string.Contains
            Assert.IsTrue(RegexHelper.IsMatch("has [invalid in text", "[invalid"));
        }

        [TestMethod]
        public void RegexHelper_IsMatch_EmptyInput_ReturnsFalse()
        {
            Assert.IsFalse(RegexHelper.IsMatch("", "sales"));
        }

        [TestMethod]
        public void RegexHelper_IsMatch_EmptyPattern_ReturnsTrue()
        {
            Assert.IsTrue(RegexHelper.IsMatch("anything", ""));
        }

        // ── RecommendationFormatter ───────────────────────────────────────────

        [TestMethod]
        public void NormalizeCapabilities_TeamMembersAlone_Kept()
        {
            var result = RecommendationFormatter.NormalizeCapabilities(new[] { "TeamMembers" });
            CollectionAssert.Contains(result, "TeamMembers");
        }

        [TestMethod]
        public void NormalizeCapabilities_TeamMembersWithHigher_Removed()
        {
            var result = RecommendationFormatter.NormalizeCapabilities(new[] { "SalesEnterprise", "TeamMembers" });
            Assert.IsFalse(result.Contains("TeamMembers"));
        }

        [TestMethod]
        public void NormalizeCapabilities_Duplicates_Deduplicated()
        {
            var result = RecommendationFormatter.NormalizeCapabilities(new[] { "SalesEnterprise", "SalesEnterprise" });
            Assert.AreEqual(1, result.Count);
        }

        [TestMethod]
        public void NormalizeCapabilities_EmptyList_ReturnsEmpty()
        {
            var result = RecommendationFormatter.NormalizeCapabilities(new string[0]);
            Assert.AreEqual(0, result.Count);
        }

        // ── TextHelper ────────────────────────────────────────────────────────

        [TestMethod]
        public void TextHelper_Shorten_LongText_Truncated()
        {
            var input = new string('a', 600);
            var result = TextHelper.Shorten(input, 500);
            Assert.AreEqual(500, result.Length);
        }

        [TestMethod]
        public void TextHelper_Shorten_ShortText_Unchanged()
        {
            var result = TextHelper.Shorten("hello", 500);
            Assert.AreEqual("hello", result);
        }

        [TestMethod]
        public void TextHelper_Shorten_Null_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, TextHelper.Shorten(null));
        }

        [TestMethod]
        public void TextHelper_Shorten_NewlinesReplaced()
        {
            var result = TextHelper.Shorten("line1\nline2\r\nline3");
            Assert.IsFalse(result.Contains("\n"));
            Assert.IsFalse(result.Contains("\r"));
        }
    }
}
