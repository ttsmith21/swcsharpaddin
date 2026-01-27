using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using NM.Core;
using NM.Core.ProblemParts;
using SolidWorks.Interop.sldworks;
using NM.SwAddin.Properties;

namespace NM.SwAddin.Processing
{
    public sealed class ProcessingCoordinator
    {
        private readonly ISldWorks _swApp;
        private readonly SolidWorksFileOperations _fileOps;
        private readonly IProcessorFactory _factory;
        private readonly ProblemPartManager _problems;
        private readonly CustomPropertiesService _propsSvc = new CustomPropertiesService();

        public ProcessingCoordinator(ISldWorks swApp, IProcessorFactory factory)
        {
            _swApp = swApp;
            _fileOps = new SolidWorksFileOperations(swApp);
            _factory = factory;
            _problems = ProblemPartManager.Instance;
        }

        public sealed class Stats
        {
            public int Total;
            public int Success;
            public int Failed;
            public int Skipped;
            public Dictionary<string, int> Used = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public TimeSpan Elapsed;
            public List<string> Errors = new List<string>();
        }

        public Stats ProcessGoodModels(IList<ModelInfo> goodModels, ProcessingOptions options)
        {
            var stats = new Stats { Total = goodModels?.Count ?? 0 };
            if (goodModels == null || goodModels.Count == 0) return stats;

            var sw = new Stopwatch(); sw.Start();
            int index = 0;
            foreach (var info in goodModels)
            {
                index++;
                IModelDoc2 model = null;
                try
                {
                    if (info.CurrentState == ModelInfo.ModelState.Processed)
                    {
                        stats.Skipped++;
                        continue;
                    }

                    model = _fileOps.OpenSWDocument(info.FilePath, silent: true, readOnly: false, configurationName: info.ConfigurationName);
                    if (model == null)
                    {
                        stats.Failed++;
                        _problems.AddProblemPart(info, "Failed to open", ProblemPartManager.ProblemCategory.FileAccess);
                        continue;
                    }

                    // Epic 0: load properties once into cache
                    _propsSvc.ReadIntoCache(model, info, includeGlobal: true, includeConfig: true);

                    var proc = _factory.DetectFor(model);
                    var result = proc.Process(model, info, options);
                    if (result.Success)
                    {
                        stats.Success++;
                        info.SetState(ModelInfo.ModelState.Processed);
                        var key = proc.Type.ToString();
                        if (!stats.Used.ContainsKey(key)) stats.Used[key] = 0;
                        stats.Used[key]++;
                    }
                    else
                    {
                        stats.Failed++;
                        stats.Errors.Add($"{info.FileName}: {result.ErrorMessage}");
                        _problems.AddProblemPart(info, result.ErrorMessage, result.ErrorCategory);
                        info.MarkProblem(result.ErrorMessage);
                    }

                    // Epic 0 legacy rule: Default/empty -> write global only; else write config only
                    _propsSvc.WritePending(model, info);

                    // Save per user choice
                    if (options != null && options.SaveChanges)
                    {
                        _fileOps.SaveSWDocument(model);
                    }
                }
                catch (Exception ex)
                {
                    stats.Failed++;
                    stats.Errors.Add($"{info.FileName}: {ex.Message}");
                    _problems.AddProblemPart(info, ex.Message, ProblemPartManager.ProblemCategory.Fatal);
                }
                finally
                {
                    // Always close
                    try { if (model != null) _fileOps.CloseSWDocument(model); } catch { }
                }
            }
            sw.Stop(); stats.Elapsed = sw.Elapsed;
            return stats;
        }

        public static string FormatSummary(Stats s)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Processing Complete");
            sb.AppendLine($"Total: {s.Total}");
            sb.AppendLine($"Success: {s.Success}");
            sb.AppendLine($"Failed: {s.Failed}");
            sb.AppendLine($"Skipped: {s.Skipped}");
            sb.AppendLine($"Elapsed: {s.Elapsed}");
            if (s.Used.Count > 0)
            {
                sb.AppendLine("Processors:");
                foreach (var kv in s.Used.OrderByDescending(k => k.Value))
                    sb.AppendLine($"  {kv.Key}: {kv.Value}");
            }
            if (s.Errors.Count > 0)
            {
                sb.AppendLine("Errors (first 10):");
                foreach (var e in s.Errors.Take(10)) sb.AppendLine("  " + e);
                if (s.Errors.Count > 10) sb.AppendLine($"  ... and {s.Errors.Count - 10} more");
            }
            return sb.ToString();
        }
    }
}
