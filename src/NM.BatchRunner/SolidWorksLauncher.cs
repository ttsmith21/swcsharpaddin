using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

namespace NM.BatchRunner
{
    /// <summary>
    /// Launches SolidWorks and connects via the Running Object Table (ROT).
    /// This is the standard pattern for stand-alone SW automation.
    /// </summary>
    public class SolidWorksLauncher : IDisposable
    {
        private Process _swProcess;
        private ISldWorks _swApp;
        private bool _disposed;

        // Path to SolidWorks 2022
        private const string SwPath = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (2)\SLDWORKS.exe";

        /// <summary>
        /// Starts SolidWorks and waits for it to be ready.
        /// </summary>
        /// <param name="timeoutSeconds">Maximum time to wait for SW to start</param>
        /// <param name="visible">Whether SW window should be visible</param>
        /// <returns>ISldWorks interface for automation</returns>
        public ISldWorks Start(int timeoutSeconds = 60, bool visible = true)
        {
            if (!System.IO.File.Exists(SwPath))
            {
                throw new System.IO.FileNotFoundException($"SolidWorks not found at: {SwPath}");
            }

            // Start SolidWorks process
            _swProcess = Process.Start(SwPath);
            Console.WriteLine($"Started SolidWorks process (PID: {_swProcess.Id})");

            // Connect via Running Object Table
            _swApp = ConnectToSwByProcessId(_swProcess.Id, timeoutSeconds);

            // Configure visibility
            _swApp.Visible = visible;

            // Wait for SolidWorks to be fully ready
            WaitForSwReady(_swApp, timeoutSeconds);

            return _swApp;
        }

        /// <summary>
        /// Connects to an already-running SolidWorks instance.
        /// </summary>
        public ISldWorks ConnectToExisting()
        {
            var swType = Type.GetTypeFromProgID("SldWorks.Application");
            _swApp = (ISldWorks)Activator.CreateInstance(swType);
            return _swApp;
        }

        private ISldWorks ConnectToSwByProcessId(int processId, int timeoutSec)
        {
            var timeout = DateTime.Now.AddSeconds(timeoutSec);
            ISldWorks swApp = null;

            while (DateTime.Now < timeout)
            {
                swApp = GetSwAppFromROT(processId);
                if (swApp != null)
                {
                    return swApp;
                }
                System.Threading.Thread.Sleep(500);
            }

            throw new TimeoutException($"Could not connect to SolidWorks process {processId} within {timeoutSec} seconds");
        }

        private void WaitForSwReady(ISldWorks swApp, int timeoutSec)
        {
            var timeout = DateTime.Now.AddSeconds(timeoutSec);

            while (DateTime.Now < timeout)
            {
                try
                {
                    // Try to access a property - if SW is ready, this will work
                    var ready = swApp.StartupProcessCompleted;
                    if (ready)
                    {
                        Console.WriteLine("SolidWorks startup completed.");
                        return;
                    }
                }
                catch
                {
                    // SW not ready yet
                }
                System.Threading.Thread.Sleep(500);
            }

            Console.WriteLine("Warning: SolidWorks startup did not complete within timeout, proceeding anyway.");
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx bc);

        private ISldWorks GetSwAppFromROT(int processId)
        {
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;
            IBindCtx bindCtx;

            int hr = GetRunningObjectTable(0, out rot);
            if (hr != 0) return null;

            rot.EnumRunning(out enumMoniker);

            hr = CreateBindCtx(0, out bindCtx);
            if (hr != 0) return null;

            IMoniker[] moniker = new IMoniker[1];
            IntPtr fetched = IntPtr.Zero;

            while (enumMoniker.Next(1, moniker, fetched) == 0)
            {
                string displayName;
                moniker[0].GetDisplayName(bindCtx, null, out displayName);

                // SolidWorks registers itself in ROT with format: SolidWorks_PID_<processId>
                if (displayName != null && displayName.Contains("SolidWorks_PID_" + processId.ToString()))
                {
                    object obj;
                    rot.GetObject(moniker[0], out obj);
                    return obj as ISldWorks;
                }
            }

            return null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Close SolidWorks gracefully
                if (_swApp != null)
                {
                    try
                    {
                        Console.WriteLine("Closing SolidWorks...");
                        _swApp.ExitApp();
                        _swApp = null;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Warning: Error closing SolidWorks: {ex.Message}");
                    }
                }

                // Wait for process to exit, then force kill if needed
                if (_swProcess != null && !_swProcess.HasExited)
                {
                    if (!_swProcess.WaitForExit(10000))
                    {
                        Console.WriteLine("SolidWorks did not exit gracefully, forcing termination...");
                        try
                        {
                            _swProcess.Kill();
                        }
                        catch { }
                    }
                    _swProcess = null;
                }
            }

            _disposed = true;
        }

        ~SolidWorksLauncher()
        {
            Dispose(false);
        }
    }
}
