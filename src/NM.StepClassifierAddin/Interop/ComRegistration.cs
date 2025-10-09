using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace NM.StepClassifierAddin.Interop
{
    /// <summary>
    /// COM registration helpers for SolidWorks AddIn registration under HKLM and AddInsStartup under HKCU.
    /// </summary>
    public static class ComRegistration
    {
        public static void RegisterAddin(Type t, string title, string description)
        {
            string clsid = t.GUID.ToString("B"); // {GUID}
            try
            {
                using (var hklm = Registry.LocalMachine)
                {
                    string addinsKey = $"SOFTWARE\\SolidWorks\\Addins\\{clsid}";
                    using (var k = hklm.CreateSubKey(addinsKey))
                    {
                        if (k == null) throw new InvalidOperationException("Failed to create Addins key");
                        k.SetValue(null, 1); // load by default
                        k.SetValue("Title", title);
                        k.SetValue("Description", description);
                    }
                }
                using (var hkcu = Registry.CurrentUser)
                {
                    string startupKey = $"Software\\SolidWorks\\AddInsStartup\\{clsid}";
                    using (var k = hkcu.CreateSubKey(startupKey))
                    {
                        if (k == null) throw new InvalidOperationException("Failed to create AddInsStartup key");
                        k.SetValue(null, 1, RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Registration failed: {ex.Message}");
                throw;
            }
        }

        public static void UnregisterAddin(Type t)
        {
            string clsid = t.GUID.ToString("B");
            try
            {
                using (var hklm = Registry.LocalMachine)
                {
                    string addinsKey = $"SOFTWARE\\SolidWorks\\Addins\\{clsid}";
                    try { hklm.DeleteSubKeyTree(addinsKey, false); } catch { }
                }
                using (var hkcu = Registry.CurrentUser)
                {
                    string startupKey = $"Software\\SolidWorks\\AddInsStartup\\{clsid}";
                    try { hkcu.DeleteSubKeyTree(startupKey, false); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Unregister failed: {ex.Message}");
            }
        }
    }
}