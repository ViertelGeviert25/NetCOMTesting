using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace COMTesting
{
    public static class ROTHelpers
    {
        // Win32-API-Aufruf zum erstellen von Bindungen
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        [DllImport("ole32.dll")]
        private static extern int CreateItemMoniker(string lpszDelim, string lpszItem, out IMoniker ppmk);

        // Win32-API-Aufruf zum lesen der ROT
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        private const int ROTFLAGS_REGISTRATIONKEEPSALIVE = 1;
        //private const int ROTFLAGS_ALLOWANYCLIENT = 2;


        [DllImport("ole32.dll", ExactSpelling = true, PreserveSig = false)]
        [Obsolete]
        private static extern UCOMIRunningObjectTable GetRunningObjectTable(int reserved);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
        [Obsolete]
        private static extern UCOMIMoniker CreateItemMoniker([In] string lpszDelim, [In] string lpszItem);


        [DllImport("ole32.dll")]
        [Obsolete]
        private static extern int CreateBindCtx(uint reserved, out UCOMIBindCtx pctx);

        /// <summary>
        /// Returns a list with display names of all currently running COM objects.
        /// </summary>
        /// <returns>List with display namesn</returns>
        public static IList<string> GetRunningComObjectNames()
        {
            // Create list of display names
            IList<string> result = new List<string>();

            // Information object of the running COM instances
            IRunningObjectTable runningObjectTable = null;

            // Moniker list
            IEnumMoniker monikerList = null;

            try
            {
                // Query Running Object Table and return nothing if no COM objects are running
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null) return null;

                // Query moniker
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                // Array für Moniker-Abfrage erzeugen
                IMoniker[] monikerContainer = new IMoniker[1];

                // Generate pointer to the number of monikers actually queried
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // Iterate through all monikers
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // Create object for binding information
                    CreateBindCtx(0, out IBindCtx bindInfo);

                    // Query the display name of the COM object via the moniker
                    monikerContainer[0].GetDisplayName(bindInfo, null, out string displayName);

                    // Dispose of binding object
                    Marshal.ReleaseComObject(bindInfo);

                    // Add display name to the listing
                    result.Add(displayName);
                }
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                // If necessary, dispose of COM references
                if (runningObjectTable != null) Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null) Marshal.ReleaseComObject(monikerList);
            }
        }

        /// <summary>
        /// Returns a reference to a running COM object based on its display name.
        /// Refer to: https://mycsharp.de/forum/threads/36340/laufende-com-objekte-abfragen?page=1
        /// </summary>
        /// <param name="objectDisplayName">Display name of a COM instance</param>
        /// <returns>Reference to COM object, or null if no COM object with the specified name is running</returns>
        public static T GetRunningComObjectByDisplayName<T>(string objectDisplayName)
        {
            // ROT-Interface
            IRunningObjectTable runningObjectTable = null;

            // Moniker- list
            IEnumMoniker monikerList = null;

            try
            {
                // Query Running Object Table and return nothing if no COM objects are running
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null) return default;

                // Query moniker
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                // Create array for moniker query
                IMoniker[] monikerContainer = new IMoniker[1];

                // Generate pointer to the number of monikers actually queried
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // Iterate through all monikers
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // Create object for binding information
                    CreateBindCtx(0, out IBindCtx bindInfo);

                    // Query the display name of the COM object via the moniker
                    monikerContainer[0].GetDisplayName(bindInfo, null, out string displayName);

                    //Console.WriteLine(displayName);

                    // Release binding object
                    Marshal.ReleaseComObject(bindInfo);

                    // If the display name matches the one you are looking for ...
                    if (displayName.ToLowerInvariant().IndexOf(objectDisplayName.ToLowerInvariant()) != -1)
                    {
                        // Query COM object via the display name
                        runningObjectTable.GetObject(monikerContainer[0], out object comInstance);

                        // Return COM object
                        return (T)comInstance;
                    }
                }
            }
            catch
            {
                // Return null
                return default;
            }
            finally
            {
                // If necessary, dispose of COM references
                if (runningObjectTable != null) Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null) Marshal.ReleaseComObject(monikerList);
            }
            // Return null
            return default;
        }

        /// <summary>
        /// Make sure to release COM object
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static T LaunchProcessAndWaitForComInstance<T>(string fileName, string arguments = "")
        {
            T app = default;

            // Start Word process without displaying the window and with the document path
            var startInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true
            };

            var proc = Process.Start(startInfo);

            // variante 1: guid, variante 2: name
            var guid = typeof(T).GUID.ToString();
            Console.WriteLine(guid);


            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                while (stopwatch.Elapsed < TimeSpan.FromSeconds(10))
                {
                    app = GetRunningComObjectByDisplayName<T>(fileName);
                    if (app != null)
                    {
                        break;
                    }
                    System.Threading.Tasks.Task.Delay(100).Wait();
                }
            }
            catch (Exception ex)
            {
                // Handle the exception appropriately
                Console.WriteLine("Error: " + ex);
            }
            finally
            {
                stopwatch.Stop();
                proc?.Dispose();
            }
            return app;
        }

        /// <summary>
        /// Unregisters the specified object from the ROT.
        /// </summary>
        /// <param name="cookie">The ROT entry to revoke.</param>
        [Obsolete]
        public static void UnregisterRunningComObject(int cookie)
        {
            UCOMIRunningObjectTable rot = null;
            UCOMIMoniker moniker = null;
            try
            {
                rot = GetRunningObjectTable(0);
                rot.Revoke(cookie);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
            finally
            {
                //Releases the COM objects
                if (moniker != null)
                    while (Marshal.ReleaseComObject(moniker) > 0) ;
                if (rot != null) while (Marshal.ReleaseComObject(rot) > 0) ;
            }
        }

        [Obsolete]
        public static int RegisterInRunningComObjectTable(object data)
        {
            int cookieValue = -1;
            UCOMIRunningObjectTable rot = null;
            UCOMIMoniker moniker = null;
            try
            {
                rot = GetRunningObjectTable(0);

                //var guid = "{" + data.GetType().GUID.ToString().ToUpper() + "}";
                var assemblyName = data.GetType().Assembly.GetName().Name;
                moniker = CreateItemMoniker("!", assemblyName); // "{0}"

                CreateBindCtx(0, out UCOMIBindCtx bindInfo);

                // what about display name????
                //var bindInfo = CreateBindCtx(0);
                //var assemblyName = data.GetType().Assembly.GetName().Name;
                //Console.WriteLine(assemblyName);
                //moniker.ParseDisplayName(bindInfo, moniker, assemblyName, out int length, out UCOMIMoniker moniker2);

                // ROTFLAGS_ALLOWANYCLIENT|
                rot.Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, data, moniker, out cookieValue);
            }
            catch
            {
                throw;
            }
            finally
            {
                //Releases the COM objects
                if (moniker != null)
                    while (Marshal.ReleaseComObject(moniker) > 0) ;
                if (rot != null) while (Marshal.ReleaseComObject(rot) > 0) ;
            }
            return cookieValue;
        }
    }
}
