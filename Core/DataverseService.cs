using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace LicenceValidator.Core
{
    public sealed class DataverseService
    {
        private readonly IOrganizationService _service;
        private readonly ToolConfig _config;
        private readonly IAuditLogger _logger;
        private readonly object _serviceLock = new object();

        public DataverseService(IOrganizationService service, ToolConfig config, IAuditLogger logger)
        {
            _service = service;
            _config = config;
            _logger = logger;
        }

        public Task<List<SystemUserRecord>> GetSystemUsersAsync(CancellationToken ct)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid", "fullname", "internalemailaddress", "domainname",
                    "azureactivedirectoryobjectid", "applicationid", "isdisabled", "accessmode"),
                Orders = { new OrderExpression("fullname", OrderType.Ascending) }
            };
            if (!_config.IncludeDisabledUsers)
                query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            if (_config.UserLimit > 0)
                query.TopCount = _config.UserLimit;
            return Task.FromResult(RetrieveAll(query).Select(ParseSystemUser).ToList());
        }

        public Task<List<AppModuleRecord>> GetAppModulesAsync(CancellationToken ct)
        {
            var query = new QueryExpression("appmodule") { ColumnSet = new ColumnSet("appmoduleid", "name", "uniquename", "isdefault") };
            return Task.FromResult(RetrieveAll(query).Select(e => new AppModuleRecord
            {
                AppModuleId = e.Id,
                Name = e.GetAttributeValue<string>("name") ?? string.Empty,
                UniqueName = e.GetAttributeValue<string>("uniquename") ?? string.Empty,
                IsDefault = e.GetAttributeValue<bool?>("isdefault")
            }).ToList());
        }

        public Task<List<AppModuleRoleLink>> GetAppModuleRolesAsync(CancellationToken ct)
        {
            var query = new QueryExpression("appmoduleroles") { ColumnSet = new ColumnSet("appmoduleid", "roleid") };
            return Task.FromResult(RetrieveAll(query).Select(e => new AppModuleRoleLink
            {
                AppModuleId = e.GetAttributeValue<EntityReference>("appmoduleid")?.Id ?? e.Id,
                RoleId = e.GetAttributeValue<EntityReference>("roleid")?.Id ?? Guid.Empty
            }).ToList());
        }

        public Task<List<EffectiveRoleRecord>> GetEffectiveRolesByAadUserAsync(Guid aadObjectId, CancellationToken ct)
        {
            var request = new OrganizationRequest("RetrieveAadUserRoles");
            request["DirectoryObjectId"] = aadObjectId;
            request.Parameters["ColumnSet"] = new ColumnSet(true);
            var response = ExecuteWithRetry(request);
            return Task.FromResult(ExtractEntityCollection(response).Select(ParseEffectiveRole).ToList());
        }

        public Task<List<RolePrivilegeInfo>> GetEffectivePrivilegesByAadUserAsync(Guid aadObjectId, CancellationToken ct)
        {
            var request = new OrganizationRequest("RetrieveAadUserPrivileges");
            request["DirectoryObjectId"] = aadObjectId;
            return Task.FromResult(ParseRolePrivileges(ExecuteWithRetry(request)));
        }

        public Task<List<EffectiveRoleRecord>> GetDirectRolesAsync(Guid systemUserId, CancellationToken ct)
        {
            var query = new QueryExpression("role") { ColumnSet = new ColumnSet("roleid", "name", "parentrootroleid") };
            var link = query.AddLink("systemuserroles", "roleid", "roleid");
            link.LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, systemUserId);
            return Task.FromResult(RetrieveAll(query).Select(e => new EffectiveRoleRecord
            {
                RoleId = e.Id,
                ParentRootRoleId = GetNullableGuid(e, "parentrootroleid"),
                Name = e.GetAttributeValue<string>("name") ?? string.Empty,
                Source = "Direct"
            }).ToList());
        }

        public Task<List<TeamRecord>> GetTeamsAsync(Guid systemUserId, CancellationToken ct)
        {
            var query = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name", "azureactivedirectoryobjectid", "membershiptype") };
            var link = query.AddLink("teammembership", "teamid", "teamid");
            link.LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, systemUserId);
            return Task.FromResult(RetrieveAll(query).Select(e => new TeamRecord
            {
                TeamId = e.Id,
                Name = e.GetAttributeValue<string>("name") ?? string.Empty,
                AzureActiveDirectoryObjectId = GetNullableGuid(e, "azureactivedirectoryobjectid"),
                MembershipType = e.GetAttributeValue<OptionSetValue>("membershiptype")?.Value
            }).ToList());
        }

        public Task<List<EffectiveRoleRecord>> GetTeamRolesAsync(Guid teamId, string teamName, CancellationToken ct)
        {
            var query = new QueryExpression("role") { ColumnSet = new ColumnSet("roleid", "name", "parentrootroleid") };
            var link = query.AddLink("teamroles", "roleid", "roleid");
            link.LinkCriteria.AddCondition("teamid", ConditionOperator.Equal, teamId);
            return Task.FromResult(RetrieveAll(query).Select(e => new EffectiveRoleRecord
            {
                RoleId = e.Id, ParentRootRoleId = GetNullableGuid(e, "parentrootroleid"),
                Name = e.GetAttributeValue<string>("name") ?? string.Empty,
                Source = "Team", SourceTeamId = teamId, SourceTeamName = teamName
            }).ToList());
        }

        public Task<List<RolePrivilegeInfo>> GetEffectivePrivilegesBySystemUserAsync(Guid systemUserId, CancellationToken ct)
        {
            var request = new OrganizationRequest("RetrieveUserPrivileges");
            request["UserId"] = systemUserId;
            return Task.FromResult(ParseRolePrivileges(ExecuteWithRetry(request)));
        }

        public Task<List<DataverseEntityMetadata>> GetUserOwnedEntityMetadataAsync(CancellationToken ct)
        {
            var request = new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = false };
            var response = (RetrieveAllEntitiesResponse)ExecuteWithRetry(request);
            return Task.FromResult(response.EntityMetadata
                .Where(m => m.OwnershipType == OwnershipTypes.UserOwned)
                .Select(ParseEntityMeta)
                .Where(x => !string.IsNullOrWhiteSpace(x.LogicalName) && !string.IsNullOrWhiteSpace(x.EntitySetName))
                .OrderBy(x => x.LogicalName, StringComparer.OrdinalIgnoreCase).ToList());
        }

        public Task<Dictionary<string, DataverseEntityMetadata>> GetEntityMetadataByLogicalNamesAsync(IEnumerable<string> logicalNames, CancellationToken ct)
        {
            var result = new Dictionary<string, DataverseEntityMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (var ln in logicalNames.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    var req = new RetrieveEntityRequest { LogicalName = ln, EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = false };
                    var resp = (RetrieveEntityResponse)ExecuteWithRetry(req);
                    var meta = ParseEntityMeta(resp.EntityMetadata);
                    if (!string.IsNullOrWhiteSpace(meta.LogicalName)) result[meta.LogicalName] = meta;
                }
                catch (Exception ex) when (ex.Message.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0
                                           || ex.Message.IndexOf("Could not find", StringComparison.OrdinalIgnoreCase) >= 0
                                           || ex.Message.IndexOf("0x80060888", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _logger.Warn("Metadata not found for entity '" + ln + "'.");
                }
                catch (Exception ex)
                {
                    // Other errors (e.g. permission denied): log but still return a stub so
                    // QueryError is set by the caller instead of silently dropping the table.
                    _logger.Warn("Metadata retrieval failed for entity '" + ln + "': " + ex.Message);
                    result[ln] = new DataverseEntityMetadata
                    {
                        LogicalName = ln,
                        EntitySetName = ln + "s",
                        PrimaryIdAttribute = ln + "id",
                        IsCustomEntity = false,
                        IsActivity = false
                    };
                }
            }
            return Task.FromResult(result);
        }
        public Task<Dictionary<Guid, long>> GetUsageCountsAsync(UsageTableProfile table, string principalAttr, string countAttr, string dateAttr, DateTime? startUtc, DateTime? endUtc, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(table.LogicalName)) return Task.FromResult(new Dictionary<Guid, long>());
            return Task.FromResult(GetUsageCountsRecursive(table, principalAttr, countAttr, dateAttr, startUtc, endUtc));
        }

        private Dictionary<Guid, long> GetUsageCountsRecursive(UsageTableProfile table, string principalAttr, string countAttr, string dateAttr, DateTime? startUtc, DateTime? endUtc)
        {
            try
            {
                var fetchXml = BuildAggregateFetchXml(table.LogicalName, principalAttr, countAttr, dateAttr, startUtc, endUtc);
                DataCollection<Entity> entities;
                lock (_serviceLock) { entities = _service.RetrieveMultiple(new FetchExpression(fetchXml)).Entities; }
                return ParseUsageCounts(entities);
            }
            catch (Exception ex) when (ShouldSplitQuery(ex) && CanSplit(startUtc, endUtc))
            {
                var mid = new DateTime(startUtc.Value.Ticks + (endUtc.Value.Ticks - startUtc.Value.Ticks) / 2, DateTimeKind.Utc);
                var left = GetUsageCountsRecursive(table, principalAttr, countAttr, dateAttr, startUtc, mid);
                var right = GetUsageCountsRecursive(table, principalAttr, countAttr, dateAttr, mid, endUtc);
                foreach (var pair in right) left[pair.Key] = left.ContainsKey(pair.Key) ? left[pair.Key] + pair.Value : pair.Value;
                return left;
            }
        }

        private static bool ShouldSplitQuery(Exception ex) => ex.ToString().IndexOf("AggregateQueryRecordLimit", StringComparison.OrdinalIgnoreCase) >= 0;
        private static bool CanSplit(DateTime? s, DateTime? e) => s.HasValue && e.HasValue && (e.Value - s.Value).TotalDays > 1;

        private static Dictionary<Guid, long> ParseUsageCounts(DataCollection<Entity> entities)
        {
            var result = new Dictionary<Guid, long>();
            foreach (var entity in entities)
            {
                Guid? principalId = null;
                if (entity.Contains("principalid"))
                {
                    var val = entity["principalid"];
                    if (val is AliasedValue av1) val = av1.Value;
                    if (val is EntityReference er) principalId = er.Id;
                    else if (val is Guid g) principalId = g;
                }
                if (!principalId.HasValue) continue;
                long count = 0;
                if (entity.Contains("recordcount"))
                {
                    var val = entity["recordcount"];
                    if (val is AliasedValue av2) val = av2.Value;
                    if (val is int i) count = i;
                    else if (val is long l) count = l;
                }
                if (count > 0) result[principalId.Value] = result.ContainsKey(principalId.Value) ? result[principalId.Value] + count : count;
            }
            return result;
        }

        // ── Internal helpers ──

        private List<Entity> RetrieveAll(QueryExpression query)
        {
            var results = new List<Entity>();
            query.PageInfo = new PagingInfo { PageNumber = 1, Count = 5000 };
            while (true)
            {
                EntityCollection response;
                lock (_serviceLock) { response = _service.RetrieveMultiple(query); }
                results.AddRange(response.Entities);
                if (!response.MoreRecords) break;
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = response.PagingCookie;
            }
            return results;
        }

        private static List<Entity> ExtractEntityCollection(OrganizationResponse response)
        {
            foreach (var kvp in response.Results)
                if (kvp.Value is EntityCollection ec) return ec.Entities.ToList();
            return new List<Entity>();
        }

        private OrganizationResponse ExecuteWithRetry(OrganizationRequest request)
        {
            int attempt = 0;
            while (true)
            {
                attempt++;
                try { lock (_serviceLock) { return _service.Execute(request); } }
                catch (Exception ex) when (attempt < _config.MaxRetryCount && IsTransient(ex))
                {
                    var delay = TimeSpan.FromSeconds(Math.Min(30, Math.Pow(2, attempt)));
                    _logger.Warn("Dataverse '" + request.RequestName + "' failed (attempt " + attempt + "). Retrying in " + delay.TotalSeconds.ToString("N1") + "s.");
                    Thread.Sleep(delay);
                }
            }
        }

        private static bool IsTransient(Exception ex)
        {
            var msg = ex.ToString();
            return msg.IndexOf("429", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("503", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("TooManyRequests", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("Throttl", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static SystemUserRecord ParseSystemUser(Entity entity)
        {
            var user = new SystemUserRecord
            {
                SystemUserId = entity.Id,
                FullName = entity.GetAttributeValue<string>("fullname"),
                InternalEmailAddress = entity.GetAttributeValue<string>("internalemailaddress"),
                DomainName = entity.GetAttributeValue<string>("domainname"),
                AzureActiveDirectoryObjectId = GetNullableGuid(entity, "azureactivedirectoryobjectid"),
                ApplicationId = GetNullableGuid(entity, "applicationid"),
                IsDisabled = entity.GetAttributeValue<bool?>("isdisabled") ?? false,
                AccessMode = entity.Contains("accessmode") ? (int?)(entity.GetAttributeValue<OptionSetValue>("accessmode")?.Value) : null
            };
            user.UserType = UserClassifier.DetermineUserType(user);
            return user;
        }

        private static EffectiveRoleRecord ParseEffectiveRole(Entity entity)
        {
            var teamId = ResolveGuid(entity.Contains("t.teamid") ? entity["t.teamid"] : null)
                         ?? ResolveGuid(entity.Contains("t_x002e_teamid") ? entity["t_x002e_teamid"] : null);
            string teamName = ResolveString(entity, "t.name") ?? ResolveString(entity, "t_x002e_name");
            return new EffectiveRoleRecord
            {
                RoleId = entity.Id != Guid.Empty ? entity.Id : GetNullableGuid(entity, "roleid") ?? Guid.Empty,
                ParentRootRoleId = GetNullableGuid(entity, "parentrootroleid") ?? GetNullableGuid(entity, "_parentrootroleid_value"),
                Name = entity.GetAttributeValue<string>("name") ?? string.Empty,
                Source = teamId.HasValue ? "Team" : "Direct",
                SourceTeamId = teamId, SourceTeamName = teamName,
                SourceTeamMembershipType = ResolveInt(entity, "t.membershiptype") ?? ResolveInt(entity, "t_x002e_membershiptype"),
                SourceTeamAzureActiveDirectoryObjectId = ResolveGuid(entity.Contains("t.azureactivedirectoryobjectid") ? entity["t.azureactivedirectoryobjectid"] : null)
            };
        }

        private static List<RolePrivilegeInfo> ParseRolePrivileges(OrganizationResponse response)
        {
            if (response.Results.TryGetValue("RolePrivileges", out var obj) && obj is EntityCollection ec)
                return ec.Entities.Select(e => new RolePrivilegeInfo
                {
                    PrivilegeName = e.GetAttributeValue<string>("PrivilegeName") ?? e.GetAttributeValue<string>("privilegename") ?? string.Empty,
                    Depth = e.GetAttributeValue<string>("Depth") ?? e.GetAttributeValue<string>("depth") ?? string.Empty,
                    PrivilegeId = e.Contains("PrivilegeId") ? e.GetAttributeValue<Guid>("PrivilegeId") : e.GetAttributeValue<Guid>("privilegeid"),
                    BusinessUnitId = e.GetAttributeValue<Guid?>("BusinessUnitId") ?? e.GetAttributeValue<Guid?>("businessunitid")
                }).ToList();
            if (response.Results.TryGetValue("RolePrivileges", out var arr) && arr is RolePrivilege[] rps)
                return rps.Select(rp => new RolePrivilegeInfo { PrivilegeId = rp.PrivilegeId, Depth = rp.Depth.ToString(), BusinessUnitId = rp.BusinessUnitId != Guid.Empty ? rp.BusinessUnitId : (Guid?)null }).ToList();
            return new List<RolePrivilegeInfo>();
        }

        private static DataverseEntityMetadata ParseEntityMeta(EntityMetadata meta) => new DataverseEntityMetadata
        {
            LogicalName = meta.LogicalName ?? string.Empty, SchemaName = meta.SchemaName ?? string.Empty,
            EntitySetName = meta.EntitySetName ?? string.Empty, PrimaryIdAttribute = meta.PrimaryIdAttribute ?? string.Empty,
            IsCustomEntity = meta.IsCustomEntity ?? false, IsActivity = meta.IsActivity ?? false
        };

        private static Guid? GetNullableGuid(Entity entity, string attr)
        {
            if (!entity.Contains(attr)) return null;
            return ResolveGuid(entity[attr]);
        }

        private static Guid? ResolveGuid(object raw)
        {
            if (raw == null) return null;
            if (raw is AliasedValue av) return ResolveGuid(av.Value);
            if (raw is Guid g) return g == Guid.Empty ? (Guid?)null : g;
            if (raw is EntityReference er) return er.Id == Guid.Empty ? (Guid?)null : er.Id;
            if (Guid.TryParse(raw.ToString(), out var parsed) && parsed != Guid.Empty) return parsed;
            return null;
        }

        private static int? ResolveInt(Entity entity, string attr)
        {
            if (!entity.Contains(attr)) return null;
            var raw = entity[attr];
            if (raw is AliasedValue av) raw = av.Value;
            if (raw is OptionSetValue osv) return osv.Value;
            if (raw is int i) return i;
            return null;
        }

        private static string ResolveString(Entity entity, string attr)
        {
            if (!entity.Contains(attr)) return null;
            var raw = entity[attr];
            if (raw is AliasedValue av) raw = av.Value;
            return raw as string ?? raw?.ToString();
        }

        private static string BuildAggregateFetchXml(string logicalName, string principalAttr, string countAttr, string dateAttr, DateTime? startUtc, DateTime? endUtc)
        {
            var sb = new StringBuilder();
            sb.Append("<fetch aggregate='true'><entity name='").Append(XmlEsc(logicalName)).Append("'>");
            sb.Append("<attribute name='").Append(XmlEsc(principalAttr)).Append("' alias='principalid' groupby='true' />");
            sb.Append("<attribute name='").Append(XmlEsc(countAttr)).Append("' alias='recordcount' aggregate='count' />");
            sb.Append("<filter type='and'>");
            sb.Append("<condition attribute='").Append(XmlEsc(principalAttr)).Append("' operator='not-null' />");
            if (!string.IsNullOrWhiteSpace(dateAttr) && startUtc.HasValue)
                sb.Append("<condition attribute='").Append(XmlEsc(dateAttr)).Append("' operator='on-or-after' value='").Append(XmlEsc(startUtc.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).Append("' />");
            if (!string.IsNullOrWhiteSpace(dateAttr) && endUtc.HasValue)
                sb.Append("<condition attribute='").Append(XmlEsc(dateAttr)).Append("' operator='on-or-before' value='").Append(XmlEsc(endUtc.Value.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).Append("' />");
            sb.Append("</filter></entity></fetch>");
            return sb.ToString();
        }

        private static string XmlEsc(string input) => input.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }
}