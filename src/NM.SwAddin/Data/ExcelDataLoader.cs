using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core;

namespace NM.SwAddin.Data
{
    // Excel COM automation loader without compile-time Office interop dependency
    // Mirrors VBA LoadAllExcelData flow and caches sheets as object[,] arrays
    public sealed class ExcelDataLoader : IDisposable
    {
        // Public tables (object[,] for compatibility)
        public object[,] OptiMaterialTable { get; private set; }
        public object[,] SSLaserSpeedsNFeeds { get; private set; }
        public object[,] CarbonLaserSpeedsNFeeds { get; private set; }
        public object[,] AluminumLaserSpeedsNFeeds { get; private set; }
        public object[,] MaxBendLengthTable { get; private set; }
        public object[,] ThicknessCheckTable { get; private set; }

        // Loaded flags
        public bool OptiMaterialLoaded { get; private set; }
        public bool SSLaserLoaded { get; private set; }
        public bool CarbonLaserLoaded { get; private set; }
        public bool AluminumLaserLoaded { get; private set; }
        public bool BendLengthLoaded { get; private set; }
        public bool ThicknessCheckLoaded { get; private set; }

        // Excel state
        private static object _excelApp; // late-bound Excel.Application
        private readonly Dictionary<string, object> _openWorkbooks = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, object[,]> _sheetCache = new Dictionary<string, object[,]>(StringComparer.OrdinalIgnoreCase);

        public bool LoadAllExcelData()
        {
            const string proc = nameof(ExcelDataLoader) + ".LoadAllExcelData";
            ErrorHandler.PushCallStack(proc);
            try
            {
                bool ok = true;
                if (!LoadMaterialData()) ok = false;
                if (!LoadStainlessSteelData()) ok = false;
                if (!LoadCarbonSteelData()) ok = false;
                if (!LoadAluminumData()) ok = false;
                if (!LoadBendData()) ok = false;
                if (!LoadThicknessCheckData()) ok = false;
                ErrorHandler.DebugLog(ok ? "[XL] All tables loaded" : "[XL] One or more tables failed");
                return ok;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception in LoadAllExcelData", ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public bool LoadMaterialData()
        {
            string file = NM.Core.Configuration.FilePaths.MaterialFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "OptiMaterial", ref tmp, ref loaded);
            if (!ok)
            {
                ok = LoadExcelData(file, "Material", ref tmp, ref loaded);
            }
            if (ok)
            {
                OptiMaterialTable = tmp; OptiMaterialLoaded = loaded;
            }
            return ok;
        }

        public bool LoadStainlessSteelData()
        {
            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "Stainless Steel", ref tmp, ref loaded);
            if (ok)
            {
                SSLaserSpeedsNFeeds = tmp; SSLaserLoaded = loaded;
            }
            // Opportunistically cache specific tabs if present
            TryCacheSheet(file, "304L");
            TryCacheSheet(file, "316L");
            TryCacheSheet(file, "309");
            TryCacheSheet(file, "2205");
            return ok;
        }

        public bool LoadCarbonSteelData()
        {
            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "Carbon Steel", ref tmp, ref loaded);
            if (ok)
            {
                CarbonLaserSpeedsNFeeds = tmp; CarbonLaserLoaded = loaded;
            }
            TryCacheSheet(file, "CS");
            return ok;
        }

        public bool LoadAluminumData()
        {
            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "Aluminum", ref tmp, ref loaded);
            if (ok)
            {
                AluminumLaserSpeedsNFeeds = tmp; AluminumLaserLoaded = loaded;
            }
            TryCacheSheet(file, "AL");
            return ok;
        }

        public bool LoadBendData()
        {
            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "Bend", ref tmp, ref loaded);
            if (ok)
            {
                MaxBendLengthTable = tmp; BendLengthLoaded = loaded;
            }
            return ok;
        }

        public bool LoadThicknessCheckData()
        {
            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] tmp = null; bool loaded = false;
            var ok = LoadExcelData(file, "ThickCheck", ref tmp, ref loaded);
            if (ok)
            {
                ThicknessCheckTable = tmp; ThicknessCheckLoaded = loaded;
            }
            return ok;
        }

        public bool VerifyAllDataLoaded()
        {
            bool ok = OptiMaterialLoaded && SSLaserLoaded && CarbonLaserLoaded && AluminumLaserLoaded && BendLengthLoaded && ThicknessCheckLoaded;
            ErrorHandler.DebugLog(ok ? "[XL] VerifyAllDataLoaded OK" : "[XL] VerifyAllDataLoaded missing tables");
            return ok;
        }

        public object[,] GetWorksheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return null;
            if (_sheetCache.TryGetValue(sheetName, out var cached)) return cached;

            string file = NM.Core.Configuration.FilePaths.LaserDataFilePath;
            object[,] temp = null; bool dummy = false;
            if (LoadExcelData(file, sheetName, ref temp, ref dummy))
            {
                _sheetCache[sheetName] = temp;
                return temp;
            }
            return null;
        }

        private bool LoadExcelData(string filePath, string sheetName, ref object[,] dataArray, ref bool isLoaded)
        {
            const string proc = nameof(ExcelDataLoader) + ".LoadExcelData";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (isLoaded && dataArray != null)
                {
                    ErrorHandler.DebugLog($"[XL] {sheetName} already loaded");
                    _sheetCache[sheetName] = dataArray;
                    return true;
                }

                var wb = OpenExcelWorkbook(filePath);
                if (wb == null)
                {
                    ErrorHandler.HandleError(proc, $"Failed to open workbook: {filePath}");
                    return false;
                }

                // Get worksheet
                object ws = null;
                try
                {
                    var sheets = wb.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, wb, null);
                    ws = sheets.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, sheets, new object[] { sheetName });
                }
                catch
                {
                    ErrorHandler.HandleError(proc, $"Worksheet not found: {sheetName}", null, ErrorHandler.LogLevel.Critical);
                    return false;
                }

                // Determine used range
                var used = ws.GetType().InvokeMember("UsedRange", System.Reflection.BindingFlags.GetProperty, null, ws, null);
                var values = used.GetType().InvokeMember("Value2", System.Reflection.BindingFlags.GetProperty, null, used, null);

                // Value2 returns object[,] for multi-cell
                if (values is object[,])
                {
                    dataArray = (object[,])values;
                    isLoaded = true;
                    _sheetCache[sheetName] = dataArray;
                    ErrorHandler.DebugLog($"[XL] Loaded {sheetName} from {filePath}");
                    return true;
                }
                else
                {
                    ErrorHandler.HandleError(proc, $"No data in sheet: {sheetName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Error loading data from {filePath} :: {sheetName}", ex, ErrorHandler.LogLevel.Critical);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private object OpenExcelWorkbook(string filePath)
        {
            const string proc = nameof(ExcelDataLoader) + ".OpenExcelWorkbook";
            try
            {
                if (string.IsNullOrWhiteSpace(filePath)) return null;
                if (_openWorkbooks.TryGetValue(filePath, out var existing)) return existing;

                if (_excelApp == null)
                {
                    var t = Type.GetTypeFromProgID("Excel.Application");
                    if (t == null)
                    {
                        ErrorHandler.HandleError(proc, "Excel is not installed (ProgID Excel.Application not found)", null, ErrorHandler.LogLevel.Critical);
                        return null;
                    }
                    _excelApp = Activator.CreateInstance(t);
                    // _excelApp.Visible = false
                    _excelApp.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, _excelApp, new object[] { false });
                }

                // Validate file exists
                if (!System.IO.File.Exists(filePath))
                {
                    ErrorHandler.HandleError(proc, $"File not found: {filePath}", null, ErrorHandler.LogLevel.Critical);
                    return null;
                }

                // Retry open
                int retries = NM.Core.Configuration.Defaults.MaxRetries;
                for (int attempt = 0; attempt < Math.Max(1, retries); attempt++)
                {
                    try
                    {
                        var workbooks = _excelApp.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, _excelApp, null);
                        var wb = workbooks.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath, true }); // ReadOnly:=true
                        _openWorkbooks[filePath] = wb;
                        return wb;
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.HandleError(proc, $"Open attempt {attempt + 1} failed: {ex.Message}", ex, ErrorHandler.LogLevel.Warning);
                        System.Threading.Thread.Sleep(200);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception creating or opening Excel", ex);
                return null;
            }
        }

        private void TryCacheSheet(string file, string name)
        {
            try
            {
                object[,] tmp = null; bool dummy = false;
                if (LoadExcelData(file, name, ref tmp, ref dummy) && tmp != null)
                {
                    _sheetCache[name] = tmp;
                }
            }
            catch { }
        }

        private static void CloseWorkbook(object wb)
        {
            try
            {
                if (wb == null) return;
                wb.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, wb, new object[] { false });
            }
            catch { }
        }

        public void ClearAllExcelData()
        {
            const string proc = nameof(ExcelDataLoader) + ".ClearAllExcelData";
            ErrorHandler.PushCallStack(proc);
            try
            {
                OptiMaterialTable = null; OptiMaterialLoaded = false;
                SSLaserSpeedsNFeeds = null; SSLaserLoaded = false;
                CarbonLaserSpeedsNFeeds = null; CarbonLaserLoaded = false;
                AluminumLaserSpeedsNFeeds = null; AluminumLaserLoaded = false;
                MaxBendLengthTable = null; BendLengthLoaded = false;
                ThicknessCheckTable = null; ThicknessCheckLoaded = false;
                _sheetCache.Clear();

                // Close workbooks
                foreach (var kv in new Dictionary<string, object>(_openWorkbooks))
                {
                    CloseWorkbook(kv.Value);
                    _openWorkbooks.Remove(kv.Key);
                }

                // Quit Excel if no workbooks remain
                if (_excelApp != null)
                {
                    try
                    {
                        var wbs = _excelApp.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, _excelApp, null);
                        int count = Convert.ToInt32(wbs.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, wbs, null));
                        if (count == 0)
                        {
                            _excelApp.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, _excelApp, null);
                            _excelApp = null;
                        }
                    }
                    catch { }
                }
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public void Dispose()
        {
            try { ClearAllExcelData(); } catch { }
        }
    }
}
