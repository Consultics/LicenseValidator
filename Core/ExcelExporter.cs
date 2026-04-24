using System;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;

namespace LicenceValidator.Core
{
    public static class ExcelExporter
    {
        private static readonly XLColor HeaderBg = XLColor.FromHtml("#4472C4");
        private static readonly XLColor HeaderFg = XLColor.White;
        private static readonly XLColor StripeBg = XLColor.FromHtml("#D9E2F3");

        public static void Export(string path, AuditResult result)
        {
            using (var wb = new XLWorkbook())
            {
                BuildUsersSheet(wb, "Users", result.UserAudits);
                BuildUsersSheet(wb, "ActionRequired", result.UserAudits
                    .Where(x => !string.Equals(x.OverallStatus, "Covered", StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(x.OverallStatus, "Special account", StringComparison.OrdinalIgnoreCase)
                                && (x.OverallStatus == null || !x.OverallStatus.StartsWith("No D365 license", StringComparison.OrdinalIgnoreCase))
                                && !(x.OverallStatus != null && x.OverallStatus.StartsWith("Licensed", StringComparison.OrdinalIgnoreCase)
                                     && string.Equals(x.CoverageStatus, "Covered", StringComparison.OrdinalIgnoreCase))
                                && !IsTeamMembersOnlyNoRecords(x))
                    .ToList());
                BuildUsersSheet(wb, "Underlicensed", result.UserAudits.Where(x => x.OverallStatus != null && x.OverallStatus.IndexOf("Underlicensed", StringComparison.OrdinalIgnoreCase) >= 0).ToList());
                BuildUsersSheet(wb, "SavingsCandidates", result.UserAudits.Where(x => x.OverallStatus != null && (x.OverallStatus.IndexOf("Overlicensed", StringComparison.OrdinalIgnoreCase) >= 0 || x.OverallStatus.IndexOf("unused", StringComparison.OrdinalIgnoreCase) >= 0 || string.Equals(x.OverallStatus, "Optimization candidate", StringComparison.OrdinalIgnoreCase))).ToList());
                wb.SaveAs(path);
            }
        }

        private static void BuildUsersSheet(XLWorkbook wb, string name, System.Collections.Generic.IReadOnlyList<UserAuditResult> audits)
        {
            var ws = wb.Worksheets.Add(name);
            var headers = new[] { "FullName", "Email", "UserType", "Disabled", "OverallStatus", "CoverageStatus", "RecommendedSku", "ActualLicenses", "Roles", "Apps", "OwnedRecords", "CreatedRecords", "ModifiedRecords", "SuggestedAction", "Notes" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = HeaderFg;
                cell.Style.Fill.BackgroundColor = HeaderBg;
            }
            ws.SheetView.FreezeRows(1);

            for (int i = 0; i < audits.Count; i++)
            {
                var a = audits[i];
                int row = i + 2;
                ws.Cell(row, 1).Value = a.User.FullName ?? string.Empty;
                ws.Cell(row, 2).Value = a.User.InternalEmailAddress ?? string.Empty;
                ws.Cell(row, 3).Value = a.User.UserType ?? string.Empty;
                ws.Cell(row, 4).Value = a.User.IsDisabled;
                ws.Cell(row, 5).Value = a.OverallStatus ?? string.Empty;
                ws.Cell(row, 6).Value = a.CoverageStatus ?? string.Empty;
                ws.Cell(row, 7).Value = a.FinalRecommendation.RecommendedSku ?? string.Empty;
                ws.Cell(row, 8).Value = string.Join(" | ", a.ActualAssignedNormalized);
                ws.Cell(row, 9).Value = string.Join(" | ", a.EffectiveRoles.Select(r => r.Source == "Direct" ? r.Name : r.Name + " [" + r.SourceTeamName + "]"));
                ws.Cell(row, 10).Value = string.Join(" | ", a.AccessibleApps.Select(ap => ap.Name));
                ws.Cell(row, 11).Value = a.Usage.OwnedRecordCount;
                ws.Cell(row, 12).Value = a.Usage.CreatedRecordCount;
                ws.Cell(row, 13).Value = a.Usage.ModifiedRecordCount;
                ws.Cell(row, 14).Value = a.SuggestedAction ?? string.Empty;
                ws.Cell(row, 15).Value = string.Join(" | ", a.Notes);

                ColorizeStatus(ws.Cell(row, 5), a.OverallStatus);
                if (i % 2 == 1)
                {
                    for (int c = 1; c <= headers.Length; c++)
                    {
                        if (c == 5) continue; // skip OverallStatus – keep its color formatting
                        ws.Cell(row, c).Style.Fill.BackgroundColor = StripeBg;
                    }
                }
            }

            if (ws.LastRowUsed() != null)
                ws.Range(1, 1, ws.LastRowUsed().RowNumber(), headers.Length).SetAutoFilter();
            ws.Columns(1, headers.Length).AdjustToContents(1, 100);
            foreach (var col in ws.Columns(1, headers.Length))
                if (col.Width > 60) col.Width = 60;
        }

        private static void ColorizeStatus(IXLCell cell, string status)
        {
            if (string.IsNullOrWhiteSpace(status)) return;
            if (status.IndexOf("Underlicensed", StringComparison.OrdinalIgnoreCase) >= 0) { cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE"); cell.Style.Font.FontColor = XLColor.FromHtml("#9C0006"); }
            else if (status.IndexOf("Overlicensed", StringComparison.OrdinalIgnoreCase) >= 0 || status.IndexOf("unused", StringComparison.OrdinalIgnoreCase) >= 0) { cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFE699"); cell.Style.Font.FontColor = XLColor.FromHtml("#9C6500"); }
            else if (string.Equals(status, "Covered", StringComparison.OrdinalIgnoreCase) || status.StartsWith("Licensed", StringComparison.OrdinalIgnoreCase)) { cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C6EFCE"); cell.Style.Font.FontColor = XLColor.FromHtml("#006100"); }
            else if (string.Equals(status, "Review", StringComparison.OrdinalIgnoreCase) || string.Equals(status, "Optimization candidate", StringComparison.OrdinalIgnoreCase)) { cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#BDD7EE"); cell.Style.Font.FontColor = XLColor.FromHtml("#1F4E79"); }
        }

        private static bool IsTeamMembersOnlyNoRecords(UserAuditResult a)
        {
            if (a.Usage.OwnedRecordCount > 0 || a.Usage.CreatedRecordCount > 0 || a.Usage.ModifiedRecordCount > 0)
                return false;
            if (a.EffectiveRoles.Count == 0) return true;
            return a.EffectiveRoles.All(r =>
                string.Equals(r.Name, "Basic User", StringComparison.OrdinalIgnoreCase)
                || r.Name.IndexOf("TeamMembers", StringComparison.OrdinalIgnoreCase) >= 0
                || r.Name.IndexOf("Team Member", StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}