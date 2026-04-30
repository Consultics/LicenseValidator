using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using LicenceValidator.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using XrmToolBox.Extensibility;

namespace LicenceValidator
{
    public partial class LicenseValidatorControl : PluginControlBase
    {
        private Settings mySettings;
        private AuditResult _lastResult;

        public LicenseValidatorControl()
        {
            InitializeComponent();
        }

        // â”€â”€ Load / Save â”€â”€

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();
                LogWarning("Settings not found => new settings file created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }

            if (mySettings.Profiles == null || mySettings.Profiles.Count == 0)
            {
                MigrateOldSettings();
            }

            RefreshProfileComboBox();
            LoadProfileToUi(mySettings.GetActiveProfile());
        }

        private void lblSectionAdvanced_Click(object sender, EventArgs e)
        {
            pnlAdvanced.Visible = !pnlAdvanced.Visible;
            lblSectionAdvanced.Text = pnlAdvanced.Visible ? "\u25BC  Advanced Settings" : "\u25B6  Advanced Settings";
        }

        private void tsbCredits_Click(object sender, EventArgs e)
        {
            using (var frm = new Form())
            {
                frm.Text = "About License Validator";
                frm.Size = new System.Drawing.Size(420, 320);
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.FormBorderStyle = FormBorderStyle.FixedDialog;
                frm.MaximizeBox = false;
                frm.MinimizeBox = false;
                frm.BackColor = System.Drawing.Color.White;

                // Logo (>C from BigImageBase64)
                var logoBytes = Convert.FromBase64String("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");
                var logoImg = System.Drawing.Image.FromStream(new System.IO.MemoryStream(logoBytes));
                var logoPic = new PictureBox
                {
                    Image = logoImg,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Size = new System.Drawing.Size(72, 72),
                    Location = new System.Drawing.Point(24, 20),
                    BackColor = System.Drawing.Color.White
                };
                frm.Controls.Add(logoPic);

                // Title
                var lblTitle = new Label
                {
                    AutoSize = true,
                    Text = "License Validator",
                    Font = new System.Drawing.Font("Segoe UI Semibold", 16F),
                    Location = new System.Drawing.Point(108, 20)
                };
                frm.Controls.Add(lblTitle);

                // Version
                var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                var lblVersion = new Label
                {
                    AutoSize = true,
                    Text = "Version " + ver.ToString(),
                    Font = new System.Drawing.Font("Segoe UI", 9.5F),
                    ForeColor = System.Drawing.Color.FromArgb(100, 100, 100),
                    Location = new System.Drawing.Point(110, 52)
                };
                frm.Controls.Add(lblVersion);

                // Author
                var lblAuthor = new Label
                {
                    AutoSize = true,
                    Text = "Author:  Martin J\u00E4ger",
                    Font = new System.Drawing.Font("Segoe UI", 10F),
                    Location = new System.Drawing.Point(110, 75)
                };
                frm.Controls.Add(lblAuthor);

                // Separator
                var sep = new Label
                {
                    Location = new System.Drawing.Point(24, 110),
                    Size = new System.Drawing.Size(360, 1),
                    BackColor = System.Drawing.Color.FromArgb(218, 218, 218)
                };
                frm.Controls.Add(sep);

                // Company
                var lblCompany = new Label
                {
                    AutoSize = true,
                    Text = "Consultics AG\nElggerstrasse 4\n8356 Ettenhausen\nSwitzerland",
                    Font = new System.Drawing.Font("Segoe UI", 9.5F),
                    Location = new System.Drawing.Point(24, 122)
                };
                frm.Controls.Add(lblCompany);

                // Website link
                var lnkWeb = new LinkLabel
                {
                    AutoSize = true,
                    Text = "www.consultics.ch",
                    Font = new System.Drawing.Font("Segoe UI", 9.5F),
                    Location = new System.Drawing.Point(24, 200),
                    LinkColor = System.Drawing.Color.FromArgb(0, 120, 212)
                };
                lnkWeb.Click += (s2, e2) => System.Diagnostics.Process.Start("https://consultics.ch");
                frm.Controls.Add(lnkWeb);

                // Copyright
                var lblCopy = new Label
                {
                    AutoSize = true,
                    Text = "\u00A9 2025 Martin J\u00E4ger, Consultics AG. All rights reserved.",
                    Font = new System.Drawing.Font("Segoe UI", 8F),
                    ForeColor = System.Drawing.Color.FromArgb(140, 140, 140),
                    Location = new System.Drawing.Point(24, 240)
                };
                frm.Controls.Add(lblCopy);

                frm.ShowDialog(this);
            }
        }

        private void MigrateOldSettings()
        {
            mySettings.Profiles = new List<SettingsProfile>
            {
                new SettingsProfile { Name = "Default" }
            };
            mySettings.ActiveProfileName = "Default";
        }

        // â”€â”€ Profile management â”€â”€

        private void RefreshProfileComboBox()
        {
            cboProfile.Items.Clear();
            foreach (var p in mySettings.Profiles)
                cboProfile.Items.Add(p.Name);
            var idx = cboProfile.Items.IndexOf(mySettings.ActiveProfileName ?? "Default");
            cboProfile.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private void cboProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            var name = cboProfile.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name)) return;
            mySettings.ActiveProfileName = name;
            LoadProfileToUi(mySettings.GetActiveProfile());
        }

        private void btnSaveProfile_Click(object sender, EventArgs e)
        {
            var name = cboProfile.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Enter a profile name.", "Save Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var profile = mySettings.Profiles.Find(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
            {
                profile = new SettingsProfile { Name = name };
                mySettings.Profiles.Add(profile);
            }
            SaveUiToProfile(profile);
            mySettings.ActiveProfileName = profile.Name;
            SaveSettings();
            RefreshProfileComboBox();
            AppendLog("Profile '" + profile.Name + "' saved.");
        }

        private void btnDeleteProfile_Click(object sender, EventArgs e)
        {
            var name = cboProfile.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name)) return;
            if (mySettings.Profiles.Count <= 1)
            {
                MessageBox.Show("Cannot delete the last profile.", "Delete Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Delete profile '" + name + "'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            mySettings.Profiles.RemoveAll(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            mySettings.ActiveProfileName = mySettings.Profiles[0].Name;
            SaveSettings();
            RefreshProfileComboBox();
            LoadProfileToUi(mySettings.GetActiveProfile());
        }

        private void LoadProfileToUi(SettingsProfile p)
        {
            if (p == null) return;
            txtTenantId.Text = p.TenantId ?? "";
            txtClientId.Text = p.ClientId ?? "";
            txtClientSecret.Text = p.ClientSecret ?? "";
            txtRulesPath.Text = p.RulesPath ?? "";
            cboAuditMode.SelectedItem = p.AuditMode ?? "RightsAndUsage";
            if (cboAuditMode.SelectedIndex < 0) cboAuditMode.SelectedItem = "RightsAndUsage";
            chkIncludeDisabledUsers.Checked = p.IncludeDisabledUsers;
            SafeSetNud(nudParallelism, 1, 20, p.MaxDegreeOfParallelism);
            SafeSetNud(nudUserLimit, 0, 99999, p.UserLimit);
            SafeSetNud(nudMaxRetry, 1, 10, p.MaxRetryCount);
            chkGraphOptional.Checked = p.GraphOptional;
            chkGraphFailOnError.Checked = p.GraphFailOnError;
            cboRecModeWithGraph.SelectedItem = p.RecommendationModeWithGraph ?? "RightsThenUsage";
            if (cboRecModeWithGraph.SelectedIndex < 0) cboRecModeWithGraph.SelectedItem = "RightsThenUsage";
            cboRecModeWithoutGraph.SelectedItem = p.RecommendationModeWithoutGraph ?? "RightsThenUsage";
            if (cboRecModeWithoutGraph.SelectedIndex < 0) cboRecModeWithoutGraph.SelectedItem = "RightsThenUsage";
            chkUsageEnabled.Checked = p.UsageEnabled;
            SafeSetNud(nudLookbackDays, 1, 9999, p.ActivityLookbackDays);
            SafeSetNud(nudOwnershipDays, 1, 9999, p.OwnershipHistoryDays);
            SafeSetNud(nudBucketDays, 1, 9999, p.BucketDays);
            chkAutoDiscoverCustom.Checked = p.AutoDiscoverCustomUserOwnedTables;
            chkIncludeStandardTables.Checked = p.IncludeStandardUserOwnedTables;
            SafeSetNud(nudMaxAutoDiscovered, 0, 9999, p.MaxAutoDiscoveredTables);
        }

        private static void SafeSetNud(System.Windows.Forms.NumericUpDown nud, int min, int max, int value)
        {
            nud.Minimum = min;
            nud.Maximum = max;
            nud.Value = Math.Max(min, Math.Min(max, value));
        }

        private void SaveUiToProfile(SettingsProfile p)
        {
            p.TenantId = txtTenantId.Text.Trim();
            p.ClientId = txtClientId.Text.Trim();
            p.ClientSecret = txtClientSecret.Text.Trim();
            p.RulesPath = txtRulesPath.Text.Trim();
            p.AuditMode = cboAuditMode.SelectedItem?.ToString() ?? "RightsAndUsage";
            p.IncludeDisabledUsers = chkIncludeDisabledUsers.Checked;
            p.MaxDegreeOfParallelism = (int)nudParallelism.Value;
            p.UserLimit = (int)nudUserLimit.Value;
            p.MaxRetryCount = (int)nudMaxRetry.Value;
            p.GraphOptional = chkGraphOptional.Checked;
            p.GraphFailOnError = chkGraphFailOnError.Checked;
            p.RecommendationModeWithGraph = cboRecModeWithGraph.SelectedItem?.ToString() ?? "RightsThenUsage";
            p.RecommendationModeWithoutGraph = cboRecModeWithoutGraph.SelectedItem?.ToString() ?? "RightsThenUsage";
            p.UsageEnabled = chkUsageEnabled.Checked;
            p.ActivityLookbackDays = (int)nudLookbackDays.Value;
            p.OwnershipHistoryDays = (int)nudOwnershipDays.Value;
            p.BucketDays = (int)nudBucketDays.Value;
            p.AutoDiscoverCustomUserOwnedTables = chkAutoDiscoverCustom.Checked;
            p.IncludeStandardUserOwnedTables = chkIncludeStandardTables.Checked;
            p.MaxAutoDiscoveredTables = (int)nudMaxAutoDiscovered.Value;
        }

        private void SaveSettings()
        {
            SaveUiToProfile(mySettings.GetActiveProfile());
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        private ToolConfig BuildToolConfig()
        {
            var p = mySettings.GetActiveProfile();
            SaveUiToProfile(p);
            return new ToolConfig
            {
                TenantId = p.TenantId,
                ClientId = p.ClientId,
                ClientSecret = p.ClientSecret,
                DataverseUrl = ConnectionDetail?.WebApplicationUrl ?? mySettings?.LastUsedOrganizationWebappUrl ?? string.Empty,
                RulesPath = p.RulesPath,
                AuditMode = p.AuditMode,
                IncludeDisabledUsers = p.IncludeDisabledUsers,
                MaxDegreeOfParallelism = p.MaxDegreeOfParallelism,
                UserLimit = p.UserLimit,
                MaxRetryCount = p.MaxRetryCount,
                GraphOptional = p.GraphOptional,
                GraphFailOnError = p.GraphFailOnError,
                RecommendationModeWithGraph = p.RecommendationModeWithGraph,
                RecommendationModeWithoutGraph = p.RecommendationModeWithoutGraph,
                UsageEnabled = p.UsageEnabled,
                ActivityLookbackDays = p.ActivityLookbackDays,
                OwnershipHistoryDays = p.OwnershipHistoryDays,
                BucketDays = p.BucketDays,
                AutoDiscoverCustomUserOwnedTables = p.AutoDiscoverCustomUserOwnedTables,
                IncludeStandardUserOwnedTables = p.IncludeStandardUserOwnedTables,
                MaxAutoDiscoveredTables = p.MaxAutoDiscoveredTables
            };
        }

        // â”€â”€ Ruleset Editor â”€â”€

        private void btnEditRuleset_Click(object sender, EventArgs e)
        {
            var json = mySettings.EmbeddedRulesetJson ?? "";

            // Load from file if specified
            if (string.IsNullOrWhiteSpace(json) && !string.IsNullOrWhiteSpace(txtRulesPath.Text) && File.Exists(txtRulesPath.Text))
            {
                json = File.ReadAllText(txtRulesPath.Text);
            }

            // Fall back to embedded default rules
            if (string.IsNullOrWhiteSpace(json))
            {
                json = LoadDefaultRulesFromResource();
            }

            // Pretty-print JSON
            json = FormatJson(json);

            using (var frm = new Form())
            {
                frm.Text = "License Rules Editor";
                frm.Size = new Size(900, 700);
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.MinimizeBox = false;

                var editor = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 10F),
                    WordWrap = false,
                    AcceptsTab = true,
                    BackColor = Color.FromArgb(30, 30, 30),
                    ForeColor = Color.FromArgb(212, 212, 212),
                    BorderStyle = BorderStyle.None,
                    DetectUrls = false,
                    HideSelection = false
                };
                SetJsonText(editor, json);

                var panel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
                var btnSave = new Button { Text = "Save to Settings", Width = 140, Location = new Point(10, 8) };
                var btnExport = new Button { Text = "Export to File...", Width = 140, Location = new Point(160, 8) };
                var btnImport = new Button { Text = "Import from File...", Width = 140, Location = new Point(310, 8) };
                var btnFormat = new Button { Text = "Format JSON", Width = 120, Location = new Point(460, 8) };
                var btnCancel = new Button { Text = "Cancel", Width = 100, Location = new Point(780, 8), DialogResult = DialogResult.Cancel };

                btnSave.Click += (s, ev) =>
                {
                    try
                    {
                        Ruleset.LoadFromJson(editor.Text);
                        mySettings.EmbeddedRulesetJson = editor.Text;
                        SaveSettings();
                        AppendLog("Ruleset saved to settings.");
                        frm.DialogResult = DialogResult.OK;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Invalid JSON:\n" + ex.Message, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                btnExport.Click += (s, ev) =>
                {
                    using (var dlg = new SaveFileDialog { Filter = "JSON|*.json", FileName = "license-rules.json" })
                    {
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            File.WriteAllText(dlg.FileName, editor.Text);
                            AppendLog("Ruleset exported: " + dlg.FileName);
                        }
                    }
                };
                btnImport.Click += (s, ev) =>
                {
                    using (var dlg = new OpenFileDialog { Filter = "JSON|*.json|All|*.*" })
                    {
                        if (dlg.ShowDialog() == DialogResult.OK)
                            SetJsonText(editor, FormatJson(File.ReadAllText(dlg.FileName)));
                    }
                };
                btnFormat.Click += (s, ev) =>
                {
                    try { SetJsonText(editor, FormatJson(editor.Text)); }
                    catch (Exception ex) { MessageBox.Show("Invalid JSON:\n" + ex.Message, "Format Error", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                };

                panel.Controls.AddRange(new Control[] { btnSave, btnExport, btnImport, btnFormat, btnCancel });
                frm.Controls.Add(editor);
                frm.Controls.Add(panel);
                frm.CancelButton = btnCancel;
                frm.ShowDialog(this);
            }
        }

        // â”€â”€ Toolbar â”€â”€

        private void tsbClose_Click(object sender, EventArgs e) => CloseTool();

        private void tsbRunAudit_Click(object sender, EventArgs e)
        {
            SaveSettings();
            ExecuteMethod(RunAuditInternal);
        }

        private void RunAuditInternal()
        {
            var config = BuildToolConfig();

            // Validate Graph settings if mode requires Graph
            if (config.GraphEnabled)
            {
                if (string.IsNullOrWhiteSpace(config.TenantId) || string.IsNullOrWhiteSpace(config.ClientId) || string.IsNullOrWhiteSpace(config.ClientSecret))
                {
                    MessageBox.Show(
                        "The selected Audit Mode requires Microsoft Graph API access,\n" +
                        "but the Graph settings (Tenant ID, Client ID, Client Secret) are not configured.\n\n" +
                        "Either fill in the Graph API settings or switch to a 'NoGraph' audit mode.",
                        "Graph Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Resolve rules: file path â†’ embedded â†’ built-in defaults
            string rulesJson = null;
            string rulesSource = null;
            if (!string.IsNullOrWhiteSpace(config.RulesPath) && File.Exists(config.RulesPath))
            {
                rulesJson = File.ReadAllText(config.RulesPath);
                rulesSource = config.RulesPath;
            }
            else if (!string.IsNullOrWhiteSpace(mySettings.EmbeddedRulesetJson))
            {
                rulesJson = mySettings.EmbeddedRulesetJson;
                rulesSource = "(embedded)";
            }
            else
            {
                rulesJson = LoadDefaultRulesFromResource();
                rulesSource = "(built-in defaults)";
            }

            txtLog.Clear();
            tsbOpenExcel.Enabled = false;
            tsbOpenExcel.BackColor = System.Drawing.Color.Transparent;
            tsbOpenExcel.ForeColor = System.Drawing.SystemColors.ControlText;
            tsbRunAudit.Enabled = false;
            _lastResult = null;

            var logger = new UiLogger(this);
            var sw = Stopwatch.StartNew();
            AppendLog("Starting audit... (Rules: " + rulesSource + ")");
            AppendLog($"[Config] AuditMode={config.AuditMode} | UsageEnabled={config.UsageEnabled} | UsageEnabledEffective={config.UsageEnabledEffective}");

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Initializing...",
                IsCancelable = true,
                Work = (worker, args) =>
                {
                    var rules = Ruleset.LoadFromJson(rulesJson);

                    using (var http = new HttpClient())
                    {
                        http.Timeout = TimeSpan.FromMinutes(5);
                        var tokenService = new OAuthTokenService(http, config);
                        var dvService = new DataverseService(Service, config, logger);
                        var graphService = new GraphService(http, tokenService, config, logger);

                        // Progress-aware engine wrapper
                        var engine = new LicenseAuditEngine(dvService, graphService, config, rules, new ProgressLogger(logger, worker, sw));
                        args.Result = engine.RunAsync(System.Threading.CancellationToken.None).GetAwaiter().GetResult();
                    }
                },
                ProgressChanged = (args) =>
                {
                    SetWorkingMessage(args.UserState?.ToString() ?? "Processing...");
                },
                PostWorkCallBack = (args) =>
                {
                    sw.Stop();
                    tsbRunAudit.Enabled = true;

                    if (args.Error != null)
                    {
                        var elapsed = sw.Elapsed;
                        AppendLog("ERROR after " + FormatDuration(elapsed) + ": " + args.Error.Message);
                        if (args.Error.InnerException != null)
                            AppendLog("  Inner: " + args.Error.InnerException.Message);
                        AppendLog("  Stack: " + args.Error.StackTrace);
                        MessageBox.Show(args.Error.Message, "Audit Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    _lastResult = (AuditResult)args.Result;
                    AppendLog("Audit completed in " + FormatDuration(sw.Elapsed) + ". " + _lastResult.UserAudits.Count + " users assessed.");
                    tsbOpenExcel.Enabled = true;
                    tsbOpenExcel.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
                    tsbOpenExcel.ForeColor = System.Drawing.Color.White;
                    PopulateResultTabs(_lastResult);
                    tabMain.SelectedTab = tabOverview;
                }
            });
        }

        private void tsbOpenExcel_Click(object sender, EventArgs e)
        {
            if (_lastResult == null)
            {
                MessageBox.Show("Please run an audit first before exporting.", "No Audit Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Export to Excel";
                dlg.Filter = "Excel Workbook|*.xlsx";
                dlg.FileName = "license-validator-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".xlsx";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExcelExporter.Export(dlg.FileName, _lastResult);
                        AppendLog("Excel exported: " + dlg.FileName);
                        MessageBox.Show("Export successful!\n" + dlg.FileName, "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        AppendLog("[ERROR] Export failed: " + ex.Message);
                        MessageBox.Show(ex.ToString(), "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnBrowseRules_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select license rules JSON file";
                dlg.Filter = "JSON Files|*.json|All Files|*.*";
                if (dlg.ShowDialog() == DialogResult.OK) txtRulesPath.Text = dlg.FileName;
            }
        }

        // â”€â”€ Result tabs â”€â”€

        private void PopulateResultTabs(AuditResult result)
        {
            PopulateOverview(result);
            PopulateUsersGrid(dgvAllUsers, result.UserAudits);
            PopulateActionRequired(dgvActionRequired, result);
            PopulateFilteredUsers(dgvUnderlicensed, result, "Underlicensed");
            PopulateFilteredUsers(dgvSavings, result, "Overlicensed", "unused", "Optimization candidate");
        }

        private void PopulateOverview(AuditResult result)
        {
            dgvOverview.Visible = false;

            // Remove any previously created overview panel
            var existing = dgvOverview.Parent.Controls["pnlOverviewSummary"];
            if (existing != null) { dgvOverview.Parent.Controls.Remove(existing); existing.Dispose(); }

            var panel = new FlowLayoutPanel
            {
                Name = "pnlOverviewSummary",
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new System.Windows.Forms.Padding(12, 12, 12, 12),
                BackColor = System.Drawing.SystemColors.Window
            };

            var m = result.Metadata;
            var audits = result.UserAudits;

            // ── Audit Info ──
            AddSectionHeader(panel, "Audit Information");
            AddField(panel, "Environment", m.DataverseUrl);
            AddField(panel, "Generated (UTC)", m.GeneratedUtc.ToString("yyyy-MM-dd HH:mm:ss"));
            AddField(panel, "Audit Mode", m.AuditMode);
            AddField(panel, "Graph Enabled", m.GraphEnabled ? "Yes" : "No");
            AddField(panel, "Usage Enabled", m.UsageEnabled ? "Yes" : "No");

            // ── User Counts ──
            AddSectionHeader(panel, "Users");
            AddField(panel, "Total Users", m.UserCount.ToString());
            AddField(panel, "Human Users", m.HumanUserCount.ToString());
            AddField(panel, "Known License State", m.KnownLicenseStateCount.ToString());

            // ── Status Breakdown ──
            AddSectionHeader(panel, "Status Breakdown");
            var groups = audits.GroupBy(x => x.OverallStatus ?? "Unknown", StringComparer.OrdinalIgnoreCase)
                               .OrderByDescending(g => g.Count()).ToList();
            foreach (var grp in groups)
                AddField(panel, grp.Key, grp.Count().ToString());

            // ── License Deficit Summary ──
            var underlicensedActive = audits.Where(a => string.Equals(a.OverallStatus, "Underlicensed (active)", StringComparison.OrdinalIgnoreCase)).ToList();
            if (underlicensedActive.Count > 0)
            {
                AddSectionHeader(panel, "Licenses to Acquire");
                AddField(panel, "Users requiring license assignment", underlicensedActive.Count.ToString());

                var byRecommendedSku = underlicensedActive
                    .GroupBy(a => a.FinalRecommendation.CommercialPattern ?? "Unknown", StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(g => g.Count())
                    .ToList();
                foreach (var grp in byRecommendedSku)
                    AddField(panel, grp.Key, grp.Count() + (grp.Count() == 1 ? " license needed" : " licenses needed"));
            }

            // ── Savings Potential ──
            var savingsCandidates = audits.Where(a =>
                (a.OverallStatus ?? "").IndexOf("Overlicensed", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(a.OverallStatus, "Licensed (unused)", StringComparison.OrdinalIgnoreCase)).ToList();
            if (savingsCandidates.Count > 0)
            {
                AddSectionHeader(panel, "Savings Potential");
                AddField(panel, "Users eligible for downgrade or removal", savingsCandidates.Count.ToString());
            }

            dgvOverview.Parent.Controls.Add(panel);
        }

        private static void AddSectionHeader(FlowLayoutPanel panel, string text)
        {
            panel.Controls.Add(new Label
            {
                Text = text,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 10f, FontStyle.Bold),
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(0, 14, 0, 4)
            });
        }

        private static void AddField(FlowLayoutPanel panel, string label, string value)
        {
            panel.Controls.Add(new Label
            {
                Text = label + ":   " + value,
                Font = SystemFonts.DefaultFont,
                AutoSize = true,
                Margin = new System.Windows.Forms.Padding(12, 2, 0, 2)
            });
        }

        private void PopulateUsersGrid(DataGridView dgv, IReadOnlyList<UserAuditResult> audits)
        {
            var dt = CreateUsersDataTable();
            foreach (var a in audits) AddUserRow(dt, a);
            dgv.DataSource = dt;
            FormatUsersGrid(dgv);
        }

        private void PopulateActionRequired(DataGridView dgv, AuditResult result)
        {
            var filtered = result.UserAudits.Where(a =>
                !string.Equals(a.OverallStatus, "Covered", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(a.OverallStatus, "Special account", StringComparison.OrdinalIgnoreCase)
                && (a.OverallStatus == null || !a.OverallStatus.StartsWith("No D365 license", StringComparison.OrdinalIgnoreCase))
                && !(a.OverallStatus != null && a.OverallStatus.StartsWith("Licensed", StringComparison.OrdinalIgnoreCase)
                     && string.Equals(a.CoverageStatus, "Covered", StringComparison.OrdinalIgnoreCase))
                && !IsTeamMembersOnlyNoRecords(a)
                || (a.OverallStatus != null && (a.OverallStatus.IndexOf("Overlicensed", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(a.OverallStatus, "Review", StringComparison.OrdinalIgnoreCase)))).ToList();
            var dt = CreateUsersDataTable();
            foreach (var a in filtered) AddUserRow(dt, a);
            dgv.DataSource = dt;
            FormatUsersGrid(dgv);
        }

        private void PopulateFilteredUsers(DataGridView dgv, AuditResult result, params string[] keywords)
        {
            var filtered = result.UserAudits.Where(a =>
                a.OverallStatus != null && keywords.Any(kw => a.OverallStatus.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
            var dt = CreateUsersDataTable();
            foreach (var a in filtered) AddUserRow(dt, a);
            dgv.DataSource = dt;
            FormatUsersGrid(dgv);
        }

        private DataTable CreateUsersDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("FullName", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("UserType", typeof(string));
            dt.Columns.Add("Disabled", typeof(bool));
            dt.Columns.Add("OverallStatus", typeof(string));
            dt.Columns.Add("CoverageStatus", typeof(string));
            dt.Columns.Add("RecommendedSku", typeof(string));
            dt.Columns.Add("ActualLicenses", typeof(string));
            dt.Columns.Add("Roles", typeof(string));
            dt.Columns.Add("Apps", typeof(string));
            dt.Columns.Add("OwnedRecords", typeof(long));
            dt.Columns.Add("CreatedRecords", typeof(long));
            dt.Columns.Add("ModifiedRecords", typeof(long));
            dt.Columns.Add("SuggestedAction", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            return dt;
        }

        private void AddUserRow(DataTable dt, UserAuditResult a)
        {
            dt.Rows.Add(
                a.User.FullName ?? "", a.User.InternalEmailAddress ?? "", a.User.UserType ?? "",
                a.User.IsDisabled, a.OverallStatus ?? "", a.CoverageStatus ?? "",
                a.FinalRecommendation.RecommendedSku ?? "",
                string.Join(" | ", a.ActualAssignedNormalized),
                string.Join(" | ", a.EffectiveRoles.Select(r => r.Name)),
                string.Join(" | ", a.AccessibleApps.Select(ap => ap.Name)),
                a.Usage.OwnedRecordCount, a.Usage.CreatedRecordCount, a.Usage.ModifiedRecordCount,
                a.SuggestedAction ?? "", string.Join(" | ", a.Notes));
        }

        private static bool IsTeamMembersOnlyNoRecords(UserAuditResult a)
        {
            if (a.Usage.OwnedRecordCount > 0 || a.Usage.CreatedRecordCount > 0 || a.Usage.ModifiedRecordCount > 0)
                return false;
            if (a.EffectiveRoles.Count == 0) return true;
            // Only has "Basic User" or "TeamMembers" style roles with no real workload
            return a.EffectiveRoles.All(r =>
                string.Equals(r.Name, "Basic User", StringComparison.OrdinalIgnoreCase)
                || r.Name.IndexOf("TeamMembers", StringComparison.OrdinalIgnoreCase) >= 0
                || r.Name.IndexOf("Team Member", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void FormatUsersGrid(DataGridView dgv)
        {
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            foreach (DataGridViewColumn col in dgv.Columns)
                if (col.Width > 300) col.Width = 300;
            dgv.CellFormatting -= Dgv_StatusCellFormatting;
            dgv.CellFormatting += Dgv_StatusCellFormatting;
        }

        private void Dgv_StatusCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var dgv = (DataGridView)sender;
            if (dgv.Columns[e.ColumnIndex].Name != "OverallStatus") return;
            var status = e.Value?.ToString();
            if (string.IsNullOrWhiteSpace(status)) return;

            if (status.IndexOf("Underlicensed", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                e.CellStyle.BackColor = Color.FromArgb(255, 199, 206);
                e.CellStyle.ForeColor = Color.FromArgb(156, 0, 6);
            }
            else if (status.IndexOf("Overlicensed", StringComparison.OrdinalIgnoreCase) >= 0 || status.IndexOf("unused", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                e.CellStyle.BackColor = Color.FromArgb(255, 230, 153);
                e.CellStyle.ForeColor = Color.FromArgb(156, 101, 0);
            }
            else if (string.Equals(status, "Covered", StringComparison.OrdinalIgnoreCase)
                     || status.StartsWith("Licensed", StringComparison.OrdinalIgnoreCase))
            {
                e.CellStyle.BackColor = Color.FromArgb(198, 239, 206);
                e.CellStyle.ForeColor = Color.FromArgb(0, 97, 0);
            }
            else if (string.Equals(status, "Review", StringComparison.OrdinalIgnoreCase)
                     || status.IndexOf("Optimization", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                e.CellStyle.BackColor = Color.FromArgb(189, 215, 238);
                e.CellStyle.ForeColor = Color.FromArgb(31, 78, 121);
            }
        }

        // â”€â”€ Helpers â”€â”€

        private void AppendLog(string message)
        {
            if (InvokeRequired) { Invoke((Action)(() => AppendLog(message))); return; }
            txtLog.AppendText("[" + DateTime.UtcNow.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }

        private static string FormatDuration(TimeSpan ts)
        {
            if (ts.TotalMinutes >= 1)
                return string.Format("{0}m {1:00}s", (int)ts.TotalMinutes, ts.Seconds);
            return string.Format("{0:0.0}s", ts.TotalSeconds);
        }

        private static string LoadDefaultRulesFromResource()
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("LicenceValidator.DefaultRules.json"))
            {
                if (stream == null) return "{ }";
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        private static string FormatJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return json;
            using (var doc = System.Text.Json.JsonDocument.Parse(json, new System.Text.Json.JsonDocumentOptions { CommentHandling = System.Text.Json.JsonCommentHandling.Skip, AllowTrailingCommas = true }))
            {
                return System.Text.Json.JsonSerializer.Serialize(doc.RootElement, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            }
        }

        private static readonly Color JsonKeyColor = Color.FromArgb(156, 220, 254);
        private static readonly Color JsonStringColor = Color.FromArgb(206, 145, 120);
        private static readonly Color JsonNumberColor = Color.FromArgb(181, 206, 168);
        private static readonly Color JsonKeywordColor = Color.FromArgb(86, 156, 214);
        private static readonly Color JsonBraceColor = Color.FromArgb(212, 212, 212);
        private static readonly Color JsonDefaultColor = Color.FromArgb(212, 212, 212);

        private static readonly Regex JsonTokenRegex = new Regex(
            @"(""(?:[^""\\]|\\.)*"")\s*:" +   // group 1: property name (followed by colon)
            @"|""(?:[^""\\]|\\.)*""" +          // string value
            @"|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?" + // number
            @"|\b(?:true|false|null)\b" +        // keywords
            @"|[{}\[\]:,]",                       // structural chars
            RegexOptions.Compiled);

        private static void SetJsonText(RichTextBox rtb, string json)
        {
            rtb.SuspendLayout();
            rtb.Text = json;
            ApplyJsonHighlighting(rtb);
            rtb.SelectionStart = 0;
            rtb.SelectionLength = 0;
            rtb.ResumeLayout();
        }

        private static void ApplyJsonHighlighting(RichTextBox rtb)
        {
            rtb.SelectAll();
            rtb.SelectionColor = JsonDefaultColor;
            rtb.SelectionStart = 0;
            rtb.SelectionLength = 0;

            var text = rtb.Text;
            // Use a simpler state-machine approach for reliability
            int i = 0;
            bool expectingValue = false;
            while (i < text.Length)
            {
                char c = text[i];

                // Skip whitespace
                if (char.IsWhiteSpace(c)) { i++; continue; }

                // String
                if (c == '"')
                {
                    int start = i;
                    i++; // skip opening quote
                    while (i < text.Length)
                    {
                        if (text[i] == '\\') { i += 2; continue; }
                        if (text[i] == '"') { i++; break; }
                        i++;
                    }
                    int len = i - start;

                    // Look ahead: is this a property name (followed by colon)?
                    int peek = i;
                    while (peek < text.Length && char.IsWhiteSpace(text[peek])) peek++;
                    bool isKey = peek < text.Length && text[peek] == ':' && !expectingValue;

                    Colorize(rtb, start, len, isKey ? JsonKeyColor : JsonStringColor);

                    if (isKey) expectingValue = false;
                    else expectingValue = false;
                    continue;
                }

                // Number
                if (c == '-' || (c >= '0' && c <= '9'))
                {
                    int start = i;
                    if (c == '-') i++;
                    while (i < text.Length && ((text[i] >= '0' && text[i] <= '9') || text[i] == '.' || text[i] == 'e' || text[i] == 'E' || text[i] == '+' || text[i] == '-'))
                    {
                        if ((text[i] == '-' || text[i] == '+') && i > start + 1 && text[i - 1] != 'e' && text[i - 1] != 'E') break;
                        i++;
                    }
                    Colorize(rtb, start, i - start, JsonNumberColor);
                    expectingValue = false;
                    continue;
                }

                // Keywords: true, false, null
                if (MatchKeyword(text, i, "true"))  { Colorize(rtb, i, 4, JsonKeywordColor); i += 4; expectingValue = false; continue; }
                if (MatchKeyword(text, i, "false")) { Colorize(rtb, i, 5, JsonKeywordColor); i += 5; expectingValue = false; continue; }
                if (MatchKeyword(text, i, "null"))  { Colorize(rtb, i, 4, JsonKeywordColor); i += 4; expectingValue = false; continue; }

                // Structural: colon signals next is a value
                if (c == ':')
                {
                    Colorize(rtb, i, 1, JsonBraceColor);
                    expectingValue = true;
                    i++; continue;
                }

                // Braces/brackets/comma
                if (c == '{' || c == '}' || c == '[' || c == ']' || c == ',')
                {
                    Colorize(rtb, i, 1, JsonBraceColor);
                    if (c == ',' || c == '{' || c == '[') expectingValue = false;
                    i++; continue;
                }

                i++;
            }
        }

        private static bool MatchKeyword(string text, int pos, string keyword)
        {
            if (pos + keyword.Length > text.Length) return false;
            for (int j = 0; j < keyword.Length; j++)
                if (text[pos + j] != keyword[j]) return false;
            // Must not be followed by a letter/digit
            if (pos + keyword.Length < text.Length && char.IsLetterOrDigit(text[pos + keyword.Length])) return false;
            return true;
        }

        private static void Colorize(RichTextBox rtb, int start, int length, Color color)
        {
            rtb.Select(start, length);
            rtb.SelectionColor = color;
        }

        private void MyPluginControl_OnCloseTool(object sender, EventArgs e) => SaveSettings();

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }

        // â”€â”€ Progress-aware logger â”€â”€

        private sealed class ProgressLogger : IAuditLogger
        {
            private readonly IAuditLogger _inner;
            private readonly System.ComponentModel.BackgroundWorker _worker;
            private readonly Stopwatch _sw;
            private int _usersDone;
            private int _usersTotal;

            public ProgressLogger(IAuditLogger inner, System.ComponentModel.BackgroundWorker worker, Stopwatch sw)
            {
                _inner = inner;
                _worker = worker;
                _sw = sw;
            }

            public void Info(string message)
            {
                _inner.Info(message);
                UpdateProgress(message);
            }

            public void Warn(string message) => _inner.Warn(message);
            public void Error(string message) => _inner.Error(message);

            private void UpdateProgress(string msg)
            {
                // "Users: 50 total, 42 to audit. Apps: 5. SKUs: 3"
                if (msg.StartsWith("Users:", StringComparison.OrdinalIgnoreCase))
                {
                    // Extract "42 to audit"
                    var idx = msg.IndexOf("to audit", StringComparison.OrdinalIgnoreCase);
                    if (idx > 0)
                    {
                        var before = msg.Substring(0, idx).Trim();
                        var lastComma = before.LastIndexOf(',');
                        if (lastComma >= 0)
                        {
                            var numPart = before.Substring(lastComma + 1).Trim();
                            int.TryParse(numPart, out _usersTotal);
                        }
                    }
                    Report("Found " + _usersTotal + " users to audit...");
                    return;
                }

                // "User 5 of 42: Max Mustermann"
                if (msg.StartsWith("User ", StringComparison.OrdinalIgnoreCase) && msg.Contains(" of "))
                {
                    // Parse "User 5 of 42"
                    try
                    {
                        var afterUser = msg.Substring(5); // "5 of 42: Max..."
                        var ofIdx = afterUser.IndexOf(" of ", StringComparison.OrdinalIgnoreCase);
                        if (ofIdx > 0)
                        {
                            int.TryParse(afterUser.Substring(0, ofIdx).Trim(), out _usersDone);
                            var afterOf = afterUser.Substring(ofIdx + 4);
                            var colonIdx = afterOf.IndexOf(':');
                            if (colonIdx > 0)
                                int.TryParse(afterOf.Substring(0, colonIdx).Trim(), out _usersTotal);
                        }
                    }
                    catch { }

                    var eta = "";
                    if (_usersDone > 2 && _usersTotal > 0)
                    {
                        var perUser = _sw.Elapsed.TotalSeconds / _usersDone;
                        var remaining = TimeSpan.FromSeconds(perUser * (_usersTotal - _usersDone));
                        eta = " | ETA: ~" + FormatDuration(remaining);
                    }
                    Report("Auditing user " + _usersDone + " of " + _usersTotal + eta);
                    return;
                }

                if (msg.IndexOf("Usage scan:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Report("Usage: " + msg.Replace("Usage scan:", "").Trim());
                    return;
                }

                if (msg.IndexOf("Audit complete", StringComparison.OrdinalIgnoreCase) >= 0)
                    Report("Finalizing...");
            }

            private void Report(string text)
            {
                var elapsed = _sw.Elapsed;
                if (elapsed.TotalSeconds > 3)
                    text += " (" + FormatDuration(elapsed) + " elapsed)";
                try { _worker.ReportProgress(0, text); } catch { }
            }

            private static string FormatDuration(TimeSpan ts)
            {
                if (ts.TotalMinutes >= 1) return string.Format("{0}m {1:00}s", (int)ts.TotalMinutes, ts.Seconds);
                return string.Format("{0:0.0}s", ts.TotalSeconds);
            }
        }

        private sealed class UiLogger : IAuditLogger
        {
            private readonly LicenseValidatorControl _ctrl;
            public UiLogger(LicenseValidatorControl ctrl) { _ctrl = ctrl; }
            public void Info(string message) { _ctrl.AppendLog(message); }
            public void Warn(string message) { _ctrl.AppendLog("[WARN] " + message); }
            public void Error(string message) { _ctrl.AppendLog("[ERROR] " + message); }
        }
    }
}

