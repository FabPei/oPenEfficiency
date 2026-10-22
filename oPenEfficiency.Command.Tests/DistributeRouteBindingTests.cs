using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Features.Alignment;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// Review round 6g measured that assertions over the command's source text can be re-expressed while every gate,
    /// both test legs and the acceptance driver stay green; review round 6h then measured that name-based rules over
    /// the built IL are not enough either, because a second type called DistributePlan with the same field names,
    /// declared in the same assembly, supplied the guarded decision and the route that reached ShapeRange.Distribute
    /// while the pinned ResolvePlan call was compiled as a discarded statement. This test therefore checks the BUILT
    /// code by identity: the rule call's declaring type AND its return type have to match the pinned types, the plan
    /// fields have to be read on exactly that return type, the route has to be read between the rule call and the COM
    /// call, and every member of the repository's own assemblies that an Execute method references has to be on the
    /// pinned list. That list lives in docs/testing/feature-acceptance/distribute-guard-pin.txt and is digest-pinned by
    /// eng/Test-FeatureAcceptanceManifests.ps1, so it cannot be extended from inside this test file; regenerate it
    /// after a legitimate change with OPENEFFICIENCY_DISTRIBUTE_GUARD_PIN=1 and commit the new digest of that gate.
    /// The opcode scan is deliberately heuristic (the call/callvirt and ldfld/stfld opcodes plus their metadata
    /// tokens): the rules are about identity, sets and order, and the scan is measured against the built command.
    /// </summary>
    [TestClass]
    public sealed class DistributeRouteBindingTests
    {
        private const string PinRelativePath = "docs/testing/feature-acceptance/distribute-guard-pin.txt";
        private const string GenerateVariable = "OPENEFFICIENCY_DISTRIBUTE_GUARD_PIN";
        private const string RouteField = "RelativeToSlide";
        private const string DecisionField = "Distribute";
        private const string PlanMethod = "ResolvePlan";
        private const string DistributeMethod = "Distribute";

        [TestMethod]
        public void DistributeCommandsApplyThePlanFromTheProductionRule()
        {
            var repositoryRoot = FindRepositoryRoot();
            Assert.IsNotNull(repositoryRoot, "The repository root (the folder holding oPenEfficiency.slnx) was not found above the test output directory.");

            var pinPath = Path.Combine(repositoryRoot, PinRelativePath.Replace('/', Path.DirectorySeparatorChar));
            if (Environment.GetEnvironmentVariable(GenerateVariable) == "1")
            {
                File.WriteAllText(pinPath, BuildPin(), new UTF8Encoding(false));
                Console.WriteLine("Wrote the distribute guard pin to " + pinPath);
                return;
            }

            Assert.IsTrue(File.Exists(pinPath), $"The distribute guard pin {PinRelativePath} is missing - a missing pin must not read as a pass.");
            var pin = Pin.Load(File.ReadAllLines(pinPath));
            var problems = new List<string>();

            foreach (var type in new[] { typeof(DistributeHorizontalFeature), typeof(DistributeVerticalFeature) })
            {
                var execute = type.GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .Single(method => method.Name == "Execute" && method.GetParameters().Length == 3);
                var label = type.Name + ".Execute";
                var calls = CallSites(execute);
                var reads = FieldSites(execute, 0x7B);
                var writes = FieldSites(execute, 0x7D);

                // (1) identity of the rule call and of the plan type it returns - a look-alike type with the same field
                // names in the same assembly is a different Type and cannot supply the values
                var planCalls = calls.Where(site => site.Member != null && site.Member.Name == PlanMethod).ToList();
                Type planType = null;
                if (planCalls.Count != 1)
                {
                    problems.Add($"{label}: expected exactly one call to {PlanMethod}, found {planCalls.Count}.");
                }
                else
                {
                    var rule = (MethodInfo)planCalls[0].Member;
                    if (!pin.MatchesRule(rule.DeclaringType))
                    {
                        problems.Add($"{label}: {PlanMethod} is called on '{Describe(rule.DeclaringType)}' instead of the pinned rule type {pin.RuleType}.");
                    }

                    planType = rule.ReturnType;
                    if (!pin.MatchesPlan(planType))
                    {
                        problems.Add($"{label}: {PlanMethod} returns '{Describe(planType)}' instead of the pinned plan type {pin.PlanType} - a second type with the same field names could supply the decision.");
                    }
                }

                var planReads = planType == null
                    ? new List<IlSite>()
                    : reads.Where(site => site.Member != null && site.Member.DeclaringType == planType).ToList();
                var lookAlikeReads = reads
                    .Where(site => site.Member != null && site.Member.DeclaringType != null
                        && site.Member.DeclaringType != planType
                        && site.Member.DeclaringType.Name == pin.PlanTypeName)
                    .ToList();
                if (lookAlikeReads.Count > 0)
                {
                    problems.Add($"{label}: reads {string.Join(", ", lookAlikeReads.Select(site => site.Member.DeclaringType.FullName + "." + site.Member.Name))} on a type that merely looks like the plan type (review round 6h).");
                }

                // (2) the two plan fields are read exactly once each, on the plan type itself, and the route is read
                // between the rule call and the COM call it feeds
                var routeReads = planReads.Where(site => site.Member.Name == RouteField).ToList();
                var decisionReads = planReads.Where(site => site.Member.Name == DecisionField).ToList();
                if (routeReads.Count != 1)
                {
                    problems.Add($"{label}: the plan's {RouteField} is read {routeReads.Count} times - exactly one read may feed the distribution.");
                }

                if (decisionReads.Count != 1)
                {
                    problems.Add($"{label}: the plan's {DecisionField} is read {decisionReads.Count} times - the guard must ask the plan exactly once.");
                }

                if (planReads.Count != 2)
                {
                    problems.Add($"{label}: the command reads {planReads.Count} plan fields ({string.Join(", ", planReads.Select(site => site.Member.Name))}) - expected exactly {DecisionField} and {RouteField}.");
                }

                var distributeCalls = calls.Where(site => site.Member != null && site.Member.Name == DistributeMethod).ToList();
                if (distributeCalls.Count == 0)
                {
                    problems.Add($"{label}: no call to {DistributeMethod} was found.");
                }
                else if (decisionReads.Count == 1 && !(decisionReads[0].Offset < distributeCalls[0].Offset))
                {
                    problems.Add($"{label}: the plan's {DecisionField} is read at IL offset {decisionReads[0].Offset}, after the {DistributeMethod} call at {distributeCalls[0].Offset} - the decision has to guard the work, not follow it (review round 6i).");
                }
                else if (planCalls.Count == 1 && routeReads.Count == 1
                    && !(planCalls[0].Offset < routeReads[0].Offset && routeReads[0].Offset < distributeCalls[0].Offset))
                {
                    problems.Add($"{label}: the route read at IL offset {routeReads[0].Offset} must sit between the rule call ({planCalls[0].Offset}) and the {DistributeMethod} call ({distributeCalls[0].Offset}).");
                }

                var planWrites = planType == null
                    ? new List<IlSite>()
                    : writes.Where(site => site.Member != null && site.Member.DeclaringType == planType).ToList();
                if (planWrites.Count > 0)
                {
                    problems.Add($"{label}: the command writes into the plan ({string.Join(", ", planWrites.Select(site => site.Member.Name))}) - only the production rule may decide it.");
                }

                // (3) every member this Execute method references that the repository itself declares has to be pinned,
                // whatever assembly it lives in and whatever the file is called
                var histogram = ReferencedRepositoryMembers(calls, reads, writes);
                if (histogram.Count == 0)
                {
                    problems.Add($"{label}: no repository member is referenced at all - the scan did not see the method.");
                }

                var method = type.FullName + ".Execute";
                var allowed = pin.MembersFor(method);
                if (allowed.Count == 0)
                {
                    problems.Add($"{label}: the pin carries no member list for {method} - regenerate it from {PinRelativePath}.");
                }

                foreach (var member in histogram.Keys.Where(member => !allowed.Contains(member)).OrderBy(member => member, StringComparer.Ordinal))
                {
                    problems.Add($"{label}: calls or reads '{member}', which is not on the pinned member list of {PinRelativePath} - a guard helper or a shadow decision can hide in any repository assembly. Regenerate the pin with {GenerateVariable}=1 after a legitimate change.");
                }

                if (!histogram.ContainsKey(pin.RuleMember))
                {
                    problems.Add($"{label}: the pinned rule member '{pin.RuleMember}' is not referenced at all.");
                }

                // the exact usage counts: a member that stops being referenced (for example the whole rollback of a
                // failed distribution) or one that is referenced once more than pinned has to fail here
                var pinnedCounts = pin.CountsFor(method);
                foreach (var pinned in pinnedCounts)
                {
                    histogram.TryGetValue(pinned.Key, out var actual);
                    if (actual != pinned.Value)
                    {
                        problems.Add($"{label}: '{pinned.Key}' is referenced {actual} times while the pin records {pinned.Value} - removing or adding a call must not stay invisible (review round 6i).");
                    }
                }

                if (pinnedCounts.Count == 0)
                {
                    problems.Add($"{label}: the pin carries no usage counts for {method} - regenerate it from {PinRelativePath}.");
                }
            }

            Assert.AreEqual(0, problems.Count, string.Join(Environment.NewLine, problems));
        }

        private static Dictionary<string, int> ReferencedRepositoryMembers(params List<IlSite>[] siteLists)
        {
            var histogram = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var site in siteLists.SelectMany(sites => sites))
            {
                if (site.Member == null || site.Member.DeclaringType == null)
                {
                    continue;
                }

                if (!IsRepositoryAssembly(site.Member.DeclaringType) || IsEmbeddedInterop(site.Member.DeclaringType))
                {
                    continue;
                }

                var key = MemberKey(site.Member);
                histogram.TryGetValue(key, out var count);
                histogram[key] = count + 1;
            }

            return histogram;
        }

        private static string BuildPin()
        {
            var builder = new StringBuilder();
            builder.AppendLine("# Generated by oPenEfficiency.Command.Tests/DistributeRouteBindingTests.cs.");
            builder.AppendLine("# Regenerate after a legitimate change with:");
            builder.AppendLine("#   OPENEFFICIENCY_DISTRIBUTE_GUARD_PIN=1 dotnet test oPenEfficiency.Command.Tests/oPenEfficiency.Command.Tests.csproj -c Release --filter DistributeRouteBinding");
            builder.AppendLine("# and update the digest in eng/Test-FeatureAcceptanceManifests.ps1 - the gate refuses an edited pin file.");
            builder.AppendLine("planType=oPenEfficiency.Services.DistributePlan|oPenEfficiency.Core");
            builder.AppendLine("ruleType=oPenEfficiency.Services.DistributeSelectionRule|oPenEfficiency.Core");
            builder.AppendLine("ruleMember=oPenEfficiency.Core|oPenEfficiency.Services.DistributeSelectionRule|" + PlanMethod);
            builder.AppendLine("routeField=" + RouteField);
            builder.AppendLine("decisionField=" + DecisionField);

            foreach (var type in new[] { typeof(DistributeHorizontalFeature), typeof(DistributeVerticalFeature) })
            {
                var execute = type.GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .Single(method => method.Name == "Execute" && method.GetParameters().Length == 3);
                builder.AppendLine("method=" + type.FullName + ".Execute");
                var histogram = ReferencedRepositoryMembers(CallSites(execute), FieldSites(execute, 0x7B), FieldSites(execute, 0x7D));
                foreach (var member in histogram.Keys.OrderBy(key => key, StringComparer.Ordinal))
                {
                    builder.AppendLine("member=" + member);
                }
                foreach (var member in histogram.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                {
                    builder.AppendLine("count=" + member.Key + "=" + member.Value);
                }
            }

            return builder.ToString().Replace("\r\n", "\n");
        }

        private static bool IsRepositoryAssembly(Type type)
        {
            var name = type.Assembly.GetName().Name;
            return name != null && name.StartsWith("oPenEfficiency", StringComparison.Ordinal);
        }

        /// <summary>
        /// True for the Office interop types the compiler embeds into the assemblies; they carry [ComImport] or live in
        /// the Microsoft.Office namespaces, and calls to them are the command's work, not a guard.
        /// </summary>
        private static bool IsEmbeddedInterop(Type type)
        {
            if (type.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComImportAttribute), false).Length > 0)
            {
                return true;
            }

            return type.Namespace != null && type.Namespace.StartsWith("Microsoft.Office", StringComparison.Ordinal);
        }

        private static string MemberKey(MemberInfo member)
        {
            return member.DeclaringType.Assembly.GetName().Name + "|" + member.DeclaringType.FullName + "|" + member.Name;
        }

        private static string Describe(Type type)
        {
            return type == null ? "?" : type.FullName + "|" + type.Assembly.GetName().Name;
        }

        private static string FindRepositoryRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "oPenEfficiency.slnx")))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return null;
        }

        private static List<IlSite> CallSites(MethodInfo method)
        {
            return Sites(method, new byte[] { 0x28, 0x6F });   // call, callvirt
        }

        private static List<IlSite> FieldSites(MethodInfo method, byte opcode)
        {
            return Sites(method, new[] { opcode });           // ldfld (0x7B), stfld (0x7D)
        }

        private static List<IlSite> Sites(MethodInfo method, byte[] opcodes)
        {
            var sites = new List<IlSite>();
            var il = method.GetMethodBody()?.GetILAsByteArray();
            if (il == null)
            {
                return sites;
            }

            for (var index = 0; index + 4 < il.Length; index++)
            {
                if (!opcodes.Contains(il[index]))
                {
                    continue;
                }

                var token = BitConverter.ToInt32(il, index + 1);
                MemberInfo member;
                try
                {
                    member = il[index] == 0x7B || il[index] == 0x7D
                        ? (MemberInfo)method.Module.ResolveField(token)
                        : method.Module.ResolveMethod(token);
                }
                catch
                {
                    continue;   // a token whose assembly is not loaded here; not part of these rules
                }

                sites.Add(new IlSite(index, member));
            }

            return sites;
        }

        private sealed class IlSite
        {
            public IlSite(int offset, MemberInfo member)
            {
                Offset = offset;
                Member = member;
            }

            public int Offset { get; }

            public MemberInfo Member { get; }
        }

        private sealed class Pin
        {
            private Pin()
            {
                Members = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
                Counts = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
            }

            public string PlanType { get; private set; }

            public string PlanTypeName { get; private set; }

            public string RuleType { get; private set; }

            public string RuleMember { get; private set; }

            public Dictionary<string, HashSet<string>> Members { get; }

            public Dictionary<string, Dictionary<string, int>> Counts { get; }

            public static Pin Load(IEnumerable<string> lines)
            {
                var pin = new Pin();
                var currentMethod = string.Empty;
                foreach (var raw in lines)
                {
                    var line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (line.StartsWith("method=", StringComparison.Ordinal))
                    {
                        currentMethod = line.Substring("method=".Length).Trim();
                        if (!pin.Members.ContainsKey(currentMethod))
                        {
                            pin.Members[currentMethod] = new HashSet<string>(StringComparer.Ordinal);
                        }

                        if (!pin.Counts.ContainsKey(currentMethod))
                        {
                            pin.Counts[currentMethod] = new Dictionary<string, int>(StringComparer.Ordinal);
                        }

                        continue;
                    }

                    if (line.StartsWith("member=", StringComparison.Ordinal))
                    {
                        if (currentMethod.Length > 0)
                        {
                            pin.Members[currentMethod].Add(line.Substring("member=".Length).Trim());
                        }

                        continue;
                    }

                    if (line.StartsWith("count=", StringComparison.Ordinal))
                    {
                        var entry = line.Substring("count=".Length);
                        var separatorAt = entry.LastIndexOf('=');
                        if (currentMethod.Length > 0 && separatorAt > 0 && int.TryParse(entry.Substring(separatorAt + 1), out var usage))
                        {
                            pin.Counts[currentMethod][entry.Substring(0, separatorAt)] = usage;
                        }

                        continue;
                    }

                    var separator = line.IndexOf('=');
                    if (separator <= 0)
                    {
                        continue;
                    }

                    var key = line.Substring(0, separator).Trim();
                    var value = line.Substring(separator + 1).Trim();
                    if (key == "planType")
                    {
                        pin.PlanType = value;
                        pin.PlanTypeName = value.Split('|')[0].Split('.').Last();
                    }
                    else if (key == "ruleType")
                    {
                        pin.RuleType = value;
                    }
                    else if (key == "ruleMember")
                    {
                        pin.RuleMember = value;
                    }
                }

                return pin;
            }

            public bool MatchesRule(Type type)
            {
                return Matches(type, RuleType);
            }

            public bool MatchesPlan(Type type)
            {
                return Matches(type, PlanType);
            }

            public HashSet<string> MembersFor(string method)
            {
                return Members.TryGetValue(method, out var members) ? members : new HashSet<string>(StringComparer.Ordinal);
            }

            public Dictionary<string, int> CountsFor(string method)
            {
                return Counts.TryGetValue(method, out var counts) ? counts : new Dictionary<string, int>(StringComparer.Ordinal);
            }

            private static bool Matches(Type type, string expected)
            {
                if (type == null || string.IsNullOrEmpty(expected))
                {
                    return false;
                }

                return (type.FullName + "|" + type.Assembly.GetName().Name) == expected;
            }
        }
    }
}
