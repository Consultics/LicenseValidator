using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace LicenceValidator.Core
{
    public sealed class LicenseAuditEngine
    {
        private readonly DataverseService _dv;
        private readonly GraphService _graph;
        private readonly ToolConfig _config;
        private readonly Ruleset _rules;
        private readonly IAuditLogger _log;

        public LicenseAuditEngine(DataverseService dv, GraphService graph, ToolConfig config, Ruleset rules, IAuditLogger log)
        { _dv = dv; _graph = graph; _config = config; _rules = rules; _log = log; }

        public async Task<AuditResult> RunAsync(CancellationToken ct = default)
        {
            var appModulesTask = _dv.GetAppModulesAsync(ct);
            var appRolesTask = _dv.GetAppModuleRolesAsync(ct);
            var usersTask = _dv.GetSystemUsersAsync(ct);
            var skusTask = _config.GraphEnabled ? _graph.GetSubscribedSkusAsync(ct) : Task.FromResult(new List<TenantSkuRecord>());
            await Task.WhenAll(appModulesTask, appRolesTask, usersTask, skusTask).ConfigureAwait(false);

            var tenantSkus = skusTask.Result.OrderBy(x => x.SkuPartNumber, StringComparer.OrdinalIgnoreCase).ToList();
            var appModules = appModulesTask.Result.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
            var appModuleRoles = appRolesTask.Result;
            var allUsers = usersTask.Result;
            var users = allUsers.Where(u => !UserClassifier.IsApplicationUser(u)).ToList();
            var skuMap = tenantSkus.GroupBy(x => x.SkuId).ToDictionary(g => g.Key, g => g.First().SkuPartNumber);

            _log.Info("Users: " + allUsers.Count + " total, " + users.Count + " to audit. Apps: " + appModules.Count + ". SKUs: " + tenantSkus.Count);

            var audits = new List<UserAuditResult>();
            var sem = new SemaphoreSlim(_config.MaxDegreeOfParallelism);
            int usersDone = 0;
            int usersTotal = users.Count;
            var tasks = users.Select(async user =>
            {
                await sem.WaitAsync(ct).ConfigureAwait(false);
                try
                {
                    var audit = await AuditUserAsync(user, appModules, appModuleRoles, skuMap, ct).ConfigureAwait(false);
                    lock (audits)
                    {
                        audits.Add(audit);
                        usersDone++;
                        _log.Info("User " + usersDone + " of " + usersTotal + ": " + (user.FullName ?? user.SystemUserId.ToString("D")));
                    }
                }
                finally { sem.Release(); }
            });
            await Task.WhenAll(tasks).ConfigureAwait(false);
            audits = audits.OrderBy(x => x.User.FullName ?? x.User.SystemUserId.ToString("D"), StringComparer.OrdinalIgnoreCase).ToList();

            var usageTables = new List<UsageTableProfile>();
            if (_config.UsageEnabledEffective)
            {
                _log.Info($"[Usage] UsageTableProfiles in rules: {_rules.UsageTableProfiles.Count}, IncludePatterns: {_rules.UsageIncludeEntityPatterns.Count}");
                usageTables = await PrepareUsageTablesAsync(ct).ConfigureAwait(false);
                var invalidTables = usageTables.Where(t => !string.IsNullOrWhiteSpace(t.QueryError)).ToList();
                if (invalidTables.Any())
                    _log.Warn($"[Usage] {invalidTables.Count} table(s) skipped: " + string.Join(", ", invalidTables.Select(t => t.LogicalName + " (" + t.QueryError + ")")));
                _log.Info("Usage tables: " + usageTables.Count(t => string.IsNullOrWhiteSpace(t.QueryError)) + " valid");
                await AnalyzeUsageAsync(audits, usageTables, ct).ConfigureAwait(false);
            }

            foreach (var audit in audits)
            {
                audit.UsageRecommendation = _config.UsageEnabledEffective
                    ? _rules.EvaluateUsage(audit.Usage)
                    : RecommendationDecision.CreateNoSignal("Usage", "Usage disabled.");
                audit.FinalRecommendation = _config.UsageEnabledEffective
                    ? RecommendationCombiner.Combine(audit.RightsRecommendation, audit.UsageRecommendation, _config.EffectiveRecommendationMode)
                    : audit.RightsRecommendation.CloneAs("Final");
                var cov = CoverageEvaluator.Evaluate(audit.User, audit.FinalRecommendation, audit.ActualAssignedNormalized, audit.Graph);
                audit.CoverageStatus = cov.Status;
                audit.Notes.AddRange(cov.Notes);
                FinalizeAssessment(audit);
            }

            _log.Info("Audit complete. " + audits.Count + " users assessed.");
            return new AuditResult
            {
                Metadata = new AuditMetadata
                {
                    GeneratedUtc = DateTime.UtcNow, DataverseUrl = _config.DataverseUrl, RulesPath = _config.RulesPath,
                    AuditMode = _config.AuditMode,
                    RecommendationModeUsed = _config.UsageEnabledEffective ? _config.EffectiveRecommendationMode : RecommendationModes.Rights,
                    GraphEnabled = _config.GraphEnabled, IncludeDisabledUsers = _config.IncludeDisabledUsers,
                    UsageEnabled = _config.UsageEnabledEffective,
                    ActivityLookbackDays = _config.ActivityLookbackDays, OwnershipHistoryDays = _config.OwnershipHistoryDays,
                    UserCount = audits.Count, HumanUserCount = audits.Count(x => !UserClassifier.IsSpecialAccount(x.User)),
                    SpecialAccountCount = audits.Count(x => UserClassifier.IsSpecialAccount(x.User)),
                    KnownLicenseStateCount = audits.Count(x => string.Equals(x.Graph.ActualLicenseState, "Known", StringComparison.OrdinalIgnoreCase)),
                    UnknownLicenseStateCount = audits.Count(x => !string.Equals(x.Graph.ActualLicenseState, "Known", StringComparison.OrdinalIgnoreCase))
                },
                TenantSkus = tenantSkus, AppModules = appModules, AppModuleRoles = appModuleRoles,
                UsageTables = usageTables, UserAudits = audits
            };
        }

        private async Task<UserAuditResult> AuditUserAsync(SystemUserRecord user, IReadOnlyList<AppModuleRecord> appMods, IReadOnlyList<AppModuleRoleLink> appRoles, IReadOnlyDictionary<Guid, string> skuMap, CancellationToken ct)
        {
            var r = new UserAuditResult { User = user, Graph = _config.GraphEnabled ? UserGraphLicenseSnapshot.CreateUnknown("NotAttempted", "Ist-Zustand konnte nicht geholt werden", true) : UserGraphLicenseSnapshot.CreateUnknown("DisabledByMode", "Ist-Zustand konnte nicht geholt werden") };
            try
            {
                bool aadAttempted = false;
                if (user.AzureActiveDirectoryObjectId.HasValue && !UserClassifier.IsApplicationUser(user))
                {
                    aadAttempted = true;
                    try
                    {
                        var allRoles = await _dv.GetEffectiveRolesByAadUserAsync(user.AzureActiveDirectoryObjectId.Value, ct).ConfigureAwait(false);
                        // Exclude roles inherited from default business unit teams (membershiptype 0)
                        r.EffectiveRoles = allRoles.Where(x => x.Source != "Team" || !x.SourceTeamMembershipType.HasValue || x.SourceTeamMembershipType.Value != 0).ToList();
                        r.EffectiveTeams = r.EffectiveRoles.Where(x => x.Source == "Team" && x.SourceTeamId.HasValue)
                            .Select(x => new TeamRecord { TeamId = x.SourceTeamId.Value, Name = x.SourceTeamName ?? string.Empty, AzureActiveDirectoryObjectId = x.SourceTeamAzureActiveDirectoryObjectId, MembershipType = x.SourceTeamMembershipType })
                            .GroupBy(x => x.TeamId).Select(g => g.First()).OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
                        r.EffectivePrivileges = await _dv.GetEffectivePrivilegesByAadUserAsync(user.AzureActiveDirectoryObjectId.Value, ct).ConfigureAwait(false);
                    }
                    catch (Exception ex) { r.Notes.Add("AAD role retrieval failed, falling back. " + TextHelper.Shorten(ex.Message)); await PopulateFallbackAsync(r, ct).ConfigureAwait(false); }
                    if (_config.GraphEnabled)
                    {
                        try { r.Graph = await _graph.GetUserLicenseSnapshotAsync(user.AzureActiveDirectoryObjectId.Value, skuMap, ct).ConfigureAwait(false); }
                        catch (Exception ex) { r.Graph = UserGraphLicenseSnapshot.CreateUnknown("Failed", "Ist-Zustand konnte nicht geholt werden", true); r.Graph.Attempted = true; r.Graph.Errors.Add("Graph failed: " + TextHelper.Shorten(ex.Message)); }
                    }
                }
                else { await PopulateFallbackAsync(r, ct).ConfigureAwait(false); }

                if (_config.GraphEnabled && !string.Equals(r.Graph.ActualLicenseState, "Known", StringComparison.OrdinalIgnoreCase) && !UserClassifier.IsApplicationUser(user))
                {
                    var upn = !string.IsNullOrWhiteSpace(user.DomainName) ? user.DomainName : user.InternalEmailAddress;
                    if (!string.IsNullOrWhiteSpace(upn))
                    {
                        try { r.Graph = await _graph.GetUserLicenseSnapshotByUpnAsync(upn, skuMap, ct).ConfigureAwait(false); }
                        catch (Exception ex) { r.Graph.Attempted = true; r.Graph.Errors.Add("Graph UPN failed: " + TextHelper.Shorten(ex.Message)); }
                    }
                }

                r.EffectiveRoles = r.EffectiveRoles.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
                r.EffectivePrivileges = r.EffectivePrivileges.GroupBy(x => x.PrivilegeName, StringComparer.OrdinalIgnoreCase).Select(g => g.First()).OrderBy(x => x.PrivilegeName, StringComparer.OrdinalIgnoreCase).ToList();
                r.AccessibleApps = ResolveAccessibleApps(appMods, appRoles, r.EffectiveRoles);
                r.ActualAssignedNormalized = _rules.NormalizeAssignedSkus(r.Graph.ActualSkuPartNumbers);

                var evidence = new UserEvidence
                {
                    RoleNames = r.EffectiveRoles.Select(x => x.Name).Where(x => !string.IsNullOrWhiteSpace(x)).ToList(),
                    AppNames = r.AccessibleApps.Select(x => x.Name).Where(x => !string.IsNullOrWhiteSpace(x)).ToList(),
                    PrivilegeNames = r.EffectivePrivileges.Select(x => x.PrivilegeName).Where(x => !string.IsNullOrWhiteSpace(x)).ToList(),
                    AssignedSkuPartNumbers = r.Graph.ActualSkuPartNumbers, IsDisabled = user.IsDisabled, AccessMode = user.AccessMode, UserType = user.UserType
                };
                r.RightsRecommendation = _rules.EvaluateRights(evidence);
                r.Notes.AddRange(r.Graph.Errors);
            }
            catch (Exception ex) { r.RightsRecommendation = RecommendationDecision.CreateReview("Rights", "Audit failed."); r.Notes.Add("Failed: " + TextHelper.Shorten(ex.ToString(), 1000)); }
            return r;
        }

        private async Task PopulateFallbackAsync(UserAuditResult r, CancellationToken ct)
        {
            var directRoles = await _dv.GetDirectRolesAsync(r.User.SystemUserId, ct).ConfigureAwait(false);
            var teams = await _dv.GetTeamsAsync(r.User.SystemUserId, ct).ConfigureAwait(false);
            // Exclude default business unit teams (membershiptype 0) – every user is auto-member,
            // and their roles would otherwise appear as the user's own roles.
            var nonDefaultTeams = teams.Where(t => t.MembershipType.HasValue && t.MembershipType.Value != 0).ToList();
            var teamRoleLists = await Task.WhenAll(nonDefaultTeams.Select(t => _dv.GetTeamRolesAsync(t.TeamId, t.Name, ct))).ConfigureAwait(false);
            var teamRoles = teamRoleLists.SelectMany(x => x).ToList();
            r.EffectiveRoles = directRoles.Concat(teamRoles).GroupBy(x => new { x.RoleId, x.Source, TeamId = x.SourceTeamId ?? Guid.Empty }).Select(g => g.First()).OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
            r.EffectiveTeams = nonDefaultTeams.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
            r.EffectivePrivileges = await _dv.GetEffectivePrivilegesBySystemUserAsync(r.User.SystemUserId, ct).ConfigureAwait(false);
        }

        private static List<AppModuleRecord> ResolveAccessibleApps(IReadOnlyList<AppModuleRecord> mods, IReadOnlyList<AppModuleRoleLink> links, IReadOnlyList<EffectiveRoleRecord> roles)
        {
            var roleIds = roles.Select(x => x.RoleId).Concat(roles.Where(x => x.ParentRootRoleId.HasValue).Select(x => x.ParentRootRoleId.Value)).ToHashSet();
            var appIds = links.Where(x => roleIds.Contains(x.RoleId)).Select(x => x.AppModuleId).ToHashSet();
            return mods.Where(x => appIds.Contains(x.AppModuleId)).OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private async Task<List<UsageTableProfile>> PrepareUsageTablesAsync(CancellationToken ct)
        {
            var profiles = _rules.UsageTableProfiles.Select(CloneProfile).ToList();
            var meta = await _dv.GetEntityMetadataByLogicalNamesAsync(profiles.Select(x => x.LogicalName), ct).ConfigureAwait(false);
            foreach (var p in profiles)
            {
                if (meta.TryGetValue(p.LogicalName, out var m)) EnrichProfile(p, m);
                else p.QueryError = "Entity '" + p.LogicalName + "' not found.";
                if (string.IsNullOrWhiteSpace(p.DisplayName)) p.DisplayName = p.LogicalName;
                if (string.IsNullOrWhiteSpace(p.PrimaryIdAttribute)) p.PrimaryIdAttribute = p.LogicalName + "id";
            }
            if (_config.AutoDiscoverCustomUserOwnedTables || _config.IncludeStandardUserOwnedTables)
            {
                var includePatterns = _rules.UsageIncludeEntityPatterns.Concat(_config.IncludeEntityLogicalNamePatterns).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
                if (includePatterns.Count > 0)
                {
                    var discovered = await _dv.GetUserOwnedEntityMetadataAsync(ct).ConfigureAwait(false);
                    var excludePatterns = _rules.UsageExcludeEntityPatterns.Concat(_config.ExcludeEntityLogicalNamePatterns).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
                    var auto = discovered
                        .Where(x => !profiles.Any(p => string.Equals(p.LogicalName, x.LogicalName, StringComparison.OrdinalIgnoreCase)))
                        .Where(x => x.IsCustomEntity ? _config.AutoDiscoverCustomUserOwnedTables : _config.IncludeStandardUserOwnedTables)
                        .Where(x => includePatterns.Any(pat => RegexHelper.IsMatch(x.LogicalName, pat)))
                        .Where(x => excludePatterns.All(pat => !RegexHelper.IsMatch(x.LogicalName, pat)))
                        .OrderBy(x => x.LogicalName, StringComparer.OrdinalIgnoreCase)
                        .Select(x => new UsageTableProfile { LogicalName = x.LogicalName, EntitySetName = x.EntitySetName, PrimaryIdAttribute = x.PrimaryIdAttribute, DisplayName = !string.IsNullOrWhiteSpace(x.SchemaName) ? x.SchemaName : x.LogicalName, IsAutoDiscovered = true, IsCustomEntity = x.IsCustomEntity, IsActivity = x.IsActivity }).ToList();
                    if (_config.MaxAutoDiscoveredTables > 0) auto = auto.Take(_config.MaxAutoDiscoveredTables).ToList();
                    profiles.AddRange(auto);
                }
            }
            return profiles.OrderBy(x => x.IsAutoDiscovered).ThenBy(x => x.LogicalName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private async Task AnalyzeUsageAsync(List<UserAuditResult> audits, List<UsageTableProfile> tables, CancellationToken ct)
        {
            var byUser = audits.ToDictionary(x => x.User.SystemUserId);
            var end = DateTime.UtcNow.Date.AddDays(1);
            var actStart = end.AddDays(-_config.ActivityLookbackDays);
            var ownStart = end.AddDays(-_config.OwnershipHistoryDays);
            var sem = new SemaphoreSlim(Math.Max(1, _config.MaxDegreeOfParallelism));
            int tableCounter = 0;
            int tableTotal = tables.Count(t => string.IsNullOrWhiteSpace(t.QueryError));
            var tasks = tables.Where(t => string.IsNullOrWhiteSpace(t.QueryError)).Select(async table =>
            {
                await sem.WaitAsync(ct).ConfigureAwait(false);
                try
                {
                    int nr = System.Threading.Interlocked.Increment(ref tableCounter);
                    _log.Info("Usage scan: " + table.LogicalName + " (Entity " + nr + " of " + tableTotal + ")");
                    var owned = table.CountOwned ? await GetCountsByWindowAsync(table, "ownerid", table.PrimaryIdAttribute, "createdon", ownStart, end, ct).ConfigureAwait(false) : new Dictionary<Guid, long>();
                    var created = table.CountCreated ? await GetCountsByWindowAsync(table, "createdby", table.PrimaryIdAttribute, "createdon", actStart, end, ct).ConfigureAwait(false) : new Dictionary<Guid, long>();
                    var modified = table.CountModified ? await GetCountsByWindowAsync(table, "modifiedby", table.PrimaryIdAttribute, "modifiedon", actStart, end, ct).ConfigureAwait(false) : new Dictionary<Guid, long>();
                    lock (byUser) { ApplyUsage(table, owned, created, modified, byUser); }
                }
                catch (Exception ex) { table.QueryError = TextHelper.Shorten(ex.Message); _log.Warn("Usage failed for " + table.LogicalName + ": " + table.QueryError); }
                finally { sem.Release(); }
            });
            await Task.WhenAll(tasks).ConfigureAwait(false);
            foreach (var a in audits) FinalizeUsage(a.Usage);
        }

        private async Task<Dictionary<Guid, long>> GetCountsByWindowAsync(UsageTableProfile table, string princ, string count, string dateAttr, DateTime start, DateTime end, CancellationToken ct)
        {
            var merged = new Dictionary<Guid, long>();
            var bucket = Math.Max(1, _config.BucketDays);
            var cursor = start;
            while (cursor < end)
            {
                var next = cursor.AddDays(bucket); if (next > end) next = end;
                var window = await _dv.GetUsageCountsAsync(table, princ, count, dateAttr, cursor, next, ct).ConfigureAwait(false);
                foreach (var p in window) merged[p.Key] = merged.ContainsKey(p.Key) ? merged[p.Key] + p.Value : p.Value;
                cursor = next;
            }
            return merged;
        }

        private static void ApplyUsage(UsageTableProfile table, IReadOnlyDictionary<Guid, long> owned, IReadOnlyDictionary<Guid, long> created, IReadOnlyDictionary<Guid, long> modified, IReadOnlyDictionary<Guid, UserAuditResult> byUser)
        {
            var ids = owned.Keys.Concat(created.Keys).Concat(modified.Keys).Distinct();
            foreach (var id in ids)
            {
                if (!byUser.TryGetValue(id, out var audit)) continue;
                long o = owned.ContainsKey(id) ? owned[id] : 0, c = created.ContainsKey(id) ? created[id] : 0, m = modified.ContainsKey(id) ? modified[id] : 0;
                if (o == 0 && c == 0 && m == 0) continue;
                audit.Usage.OwnedRecordCount += o; audit.Usage.CreatedRecordCount += c; audit.Usage.ModifiedRecordCount += m;
                var delta = o + c + m;
                if (table.CountsAsBusinessSignal) audit.Usage.BusinessSignalRecordCount += delta;
                if (!string.IsNullOrWhiteSpace(table.Capability))
                    audit.Usage.CapabilitySignalCounts[table.Capability] = audit.Usage.CapabilitySignalCounts.ContainsKey(table.Capability) ? audit.Usage.CapabilitySignalCounts[table.Capability] + delta : delta;
                audit.Usage.TableSignals.Add(new UserUsageTableSignal { TableLogicalName = table.LogicalName, TableDisplayName = string.IsNullOrWhiteSpace(table.DisplayName) ? table.LogicalName : table.DisplayName, Capability = table.Capability, OwnedCount = o, CreatedCount = c, ModifiedCount = m });
            }
        }

        private static void FinalizeUsage(UserUsageSummary u)
        {
            u.DetectedCapabilities = u.CapabilitySignalCounts.OrderByDescending(x => x.Value).Select(x => x.Key).ToList();
            if (!u.HasAnyBusinessDataSignal) { u.Status = "No business-data signal"; return; }
            u.Status = u.DetectedCapabilities.Count > 0 ? "Mapped workload: " + string.Join(" + ", u.DetectedCapabilities.Select(RecommendationFormatter.CapabilityToLabel)) : "Only generic signal";
            if (!u.HasRecentCreateOrModify && u.OwnedRecordCount > 0) u.Notes.Add("Ownership signal only.");
        }

        private void FinalizeAssessment(UserAuditResult a)
        {
            if (UserClassifier.IsSpecialAccount(a.User)) { a.OverallStatus = "Special account"; a.SuggestedAction = "Review separately."; return; }
            if (!string.Equals(a.Graph.ActualLicenseState, "Known", StringComparison.OrdinalIgnoreCase)) { a.OverallStatus = "Recommendation only"; a.SuggestedAction = "Validate current state and compare to: " + a.FinalRecommendation.CommercialPattern; return; }
            if (string.Equals(a.CoverageStatus, "Underlicensed", StringComparison.OrdinalIgnoreCase))
            {
                if (_config.UsageEnabledEffective && !a.Usage.HasAnyBusinessDataSignal) { a.OverallStatus = "Underlicensed (unused)"; a.SuggestedAction = "Remove rights or adjust roles."; }
                else if (_config.UsageEnabledEffective && !a.Usage.HasRecentCreateOrModify && a.Usage.OwnedRecordCount > 0) { a.OverallStatus = "Underlicensed (stale usage)"; a.SuggestedAction = "Remove rights or adjust roles – only ownership signal, no recent activity."; }
                else { a.OverallStatus = "Underlicensed (active)"; a.SuggestedAction = "Assign or upgrade license to: " + a.FinalRecommendation.CommercialPattern; }
                return;
            }
            if (string.Equals(a.CoverageStatus, "Review", StringComparison.OrdinalIgnoreCase))
            {
                var noLicense = a.ActualAssignedNormalized.Count == 0;
                var noRec = a.FinalRecommendation.Capabilities.Count == 0 || a.FinalRecommendation.Capabilities.All(c => string.Equals(c, "NeedsReview", StringComparison.OrdinalIgnoreCase));
                var noUsage = !a.Usage.HasAnyBusinessDataSignal;
                if (noLicense && noRec && noUsage) { a.OverallStatus = "No D365 license needed"; a.SuggestedAction = "No action required."; return; }
                if (noLicense && noRec && !noUsage) { a.OverallStatus = "No D365 license (has usage)"; a.SuggestedAction = "Verify if D365 license is needed."; return; }
                a.OverallStatus = "Review"; a.SuggestedAction = "Review manually."; return;
            }
            // Covered path - check for optimization
            var hints = new List<string>();
            if (a.FinalRecommendation.Capabilities.Count == 1 && a.FinalRecommendation.Capabilities.Any(c => string.Equals(c, "TeamMembers", StringComparison.OrdinalIgnoreCase)) && CoverageEvaluator.HasAnyFullBase(a.ActualAssignedNormalized))
                hints.Add("Recommendation is Team Members but full base SKU assigned.");
            if (_config.UsageEnabledEffective && a.ActualAssignedNormalized.Count > 0 && !a.Usage.HasAnyBusinessDataSignal)
                hints.Add("No business-data signal.");
            a.OptimizationStatus = string.Join(" | ", hints);
            if (hints.Count > 0)
            {
                bool overlic = hints.Any(h => h.IndexOf("full base", StringComparison.OrdinalIgnoreCase) >= 0);
                bool noUse = hints.Any(h => h.IndexOf("No business-data", StringComparison.OrdinalIgnoreCase) >= 0);
                if (overlic && noUse) { a.OverallStatus = "Overlicensed (unused)"; a.SuggestedAction = "Review license / consider downgrade or removal."; }
                else if (overlic) { a.OverallStatus = "Overlicensed"; a.SuggestedAction = "Review license / consider downgrade to: " + a.FinalRecommendation.CommercialPattern; }
                else if (noUse) { a.OverallStatus = "Licensed (unused)"; a.SuggestedAction = "Review license / consider downgrade."; }
                else { a.OverallStatus = "Optimization candidate"; a.SuggestedAction = "Review with the business."; }
            }
            else { a.OverallStatus = "Covered"; a.SuggestedAction = "No action required."; }
        }

        private static UsageTableProfile CloneProfile(UsageTableProfile s) => new UsageTableProfile
        {
            LogicalName = s.LogicalName, EntitySetName = s.EntitySetName, PrimaryIdAttribute = s.PrimaryIdAttribute,
            DisplayName = s.DisplayName, Capability = s.Capability, CountOwned = s.CountOwned, CountCreated = s.CountCreated,
            CountModified = s.CountModified, CountsAsBusinessSignal = s.CountsAsBusinessSignal,
            IsAutoDiscovered = s.IsAutoDiscovered, IsCustomEntity = s.IsCustomEntity, IsActivity = s.IsActivity, QueryError = s.QueryError
        };

        private static void EnrichProfile(UsageTableProfile p, DataverseEntityMetadata m)
        {
            if (!string.IsNullOrWhiteSpace(m.EntitySetName)) p.EntitySetName = m.EntitySetName;
            if (!string.IsNullOrWhiteSpace(m.PrimaryIdAttribute)) p.PrimaryIdAttribute = m.PrimaryIdAttribute;
            if (string.IsNullOrWhiteSpace(p.DisplayName)) p.DisplayName = !string.IsNullOrWhiteSpace(m.SchemaName) ? m.SchemaName : m.LogicalName;
            p.IsCustomEntity = m.IsCustomEntity; p.IsActivity = m.IsActivity;
        }
    }
}