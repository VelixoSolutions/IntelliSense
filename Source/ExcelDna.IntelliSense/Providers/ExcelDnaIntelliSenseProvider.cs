using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace ExcelDna.IntelliSense
{
    // Provides IntelliSense info for all Excel-DNA based .xll add-ins, using the built-in RegistrationInfo helper function.
    internal class ExcelDnaIntelliSenseProvider : IIntelliSenseProvider
    {
        private class XllRegistrationInfo
        {
            private readonly string _xllPath;
            private bool _regInfoNotAvailable = false;  // Set to true if we know for sure that reginfo is #N/A
            private double _version = -1;               // Version indicator to enumerate from scratch
            private object[,] _regInfo = null;          // Default value

            public XllRegistrationInfo(string xllPath)
            {
                _xllPath = xllPath;
            }

            // Called on the main thread in a macro context
            public void Refresh()
            {
                if (_regInfoNotAvailable)
                {
                    return;
                }

                var regInfoResponse = ExcelIntegration.GetRegistrationInfo(_xllPath, _version);

                if (regInfoResponse.Equals(ExcelError.ExcelErrorNA))
                {
                    _regInfoNotAvailable = true;
                    Logger.Provider.Info($"XllRegistrationInfo not available for {_xllPath}");
                    return;
                }

                if (regInfoResponse == null || regInfoResponse.Equals(ExcelError.ExcelErrorNum))
                {
                    // no update - versions match
                    return;
                }

                _regInfo = regInfoResponse as object[,];
                if (_regInfo != null)
                {
                    Debug.Assert((string)_regInfo[0, 0] == _xllPath);
                    _version = (double)_regInfo[0, 1];
                }
            }

            // Not in macro context - don't call Excel, could be any thread.
            public IEnumerable<FunctionInfo> GetFunctionInfos()
            {
                // to avoid worries about locking and this being updated from another thread, we take a copy of _regInfo
                var regInfo = _regInfo;
                /*
                    result[0, 0] = XlAddIn.PathXll;
                    result[0, 1] = registrationInfoVersion;
                 */

                if (regInfo == null)
                {
                    yield break;
                }

                var regInfoCount = regInfo.GetLength(0);
                Logger.Provider.Verbose($"XllRegistrationInfo for {_xllPath}: {regInfoCount} registrations");
                for (var i = 0; i < regInfoCount; i++)
                {
                    if (regInfo[i, 0] is ExcelEmpty)
                    {
                        var functionName = regInfo[i, 3] as string;
                        var helpTopic = regInfo[i, 8] as string;
                        var description = regInfo[i, 9] as string;

                        var argumentStr = regInfo[i, 4] as string;
                        var argumentNames = string.IsNullOrEmpty(argumentStr) ? Array.Empty<string>() : argumentStr.Split(',');

                        var argumentInfos = new List<FunctionInfo.ArgumentInfo>();
                        for (var j = 0; j < argumentNames.Length; j++)
                        {
                            argumentInfos.Add(new FunctionInfo.ArgumentInfo
                            {
                                Name = argumentNames[j],
                                Description = regInfo[i, j + 10] as string
                            });
                        }

                        yield return new FunctionInfo
                        {
                            Name = functionName,
                            Description = description,
                            HelpTopic = helpTopic,
                            ArgumentList = argumentInfos,
                            SourcePath = _xllPath
                        };
                    }
                }
            }
        }

        private readonly SynchronizationContext _syncContextMain; // Main thread, not macro context
        private readonly ExcelSynchronizationContext _syncContextExcel; // Proper macro context
        private readonly Dictionary<string, XllRegistrationInfo> _xllRegistrationInfos = new Dictionary<string, XllRegistrationInfo>();
        private readonly LoaderNotification _loaderNotification;
        private bool _isDirty;
        private readonly SendOrPostCallback _processLoadNotification;
        public event EventHandler Invalidate;

        public ExcelDnaIntelliSenseProvider(SynchronizationContext syncContextMain)
        {
            _loaderNotification = new LoaderNotification();
            _loaderNotification.LoadNotification += loaderNotification_LoadNotification;
            _syncContextMain = syncContextMain;
            _syncContextExcel = new ExcelSynchronizationContext();
            _processLoadNotification = ProcessLoadNotification;
        }

        #region IIntelliSenseProvider implementation

        // Must be called on the main Excel thread
        public void Initialize()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Initialize");

            lock (_xllRegistrationInfos)
            {
                Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Initialize - inside lock");
                foreach (var xllPath in GetLoadedXllPaths())
                {
                    if (!_xllRegistrationInfos.ContainsKey(xllPath))
                    {
                        Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.Initialize: Adding XllRegistrationInfo for {xllPath}");
                        var regInfo = new XllRegistrationInfo(xllPath);
                        _xllRegistrationInfos[xllPath] = regInfo;

                        regInfo.Refresh();
                    }
                }
            }
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Initialize - after lock");
        }

        // Must be called on the main Excel thread
        public void Refresh()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Refresh");
            lock (_xllRegistrationInfos)
            {
                Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Refresh - inside lock");
                foreach (XllRegistrationInfo regInfo in _xllRegistrationInfos.Values)
                {
                    regInfo.Refresh();
                }
                _isDirty = false;
            }
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Refresh - after lock");
        }

        // May be called from any thread
        public IList<FunctionInfo> GetFunctionInfos()
        {
            IList<FunctionInfo> excelDnaInfos;
            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos");
            lock (_xllRegistrationInfos)
            {
                Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos - inside lock");
                excelDnaInfos = _xllRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
            }
            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos - after lock");
            foreach (FunctionInfo info in excelDnaInfos)
            {
                Logger.Provider.Verbose($"\t{info.Name}({info.ArgumentList.Count}) - {info.Description} ");
            }

            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos End");

            return excelDnaInfos;
        }

        #endregion

        // DANGER: Still subject to LoaderLock warning ...
        // TODO: Consider Load/Unload done when AddIns is enumerated...
        private void loaderNotification_LoadNotification(object sender, LoaderNotification.NotificationEventArgs e)
        {
            Debug.Print($"@>@>@>@> LoadNotification: {e.Reason} - {e.FullDllName}");
            if (e.FullDllName.EndsWith(".xll", StringComparison.OrdinalIgnoreCase))
            {
                _syncContextMain.Post(_processLoadNotification, e);
            }
        }

        // Runs on the main thread, but not in a macro context 
        // WARNING: The sequence of calls here, between queued 
        //          instances of ProcessLoadNotification, Refresh and OnInvalidate
        //          is quite fragile.
        private void ProcessLoadNotification(object state)
        {
            var notification = (LoaderNotification.NotificationEventArgs)state;
            var xllPath = notification.FullDllName;

            Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.ProcessLoadNotification {notification}, {xllPath}");
            lock (_xllRegistrationInfos)
            {
                Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.ProcessLoadNotification - inside lock");
                if (!_xllRegistrationInfos.TryGetValue(xllPath, out XllRegistrationInfo regInfo))
                {
                    if (notification.Reason == LoaderNotification.Reason.Loaded)
                    {
                        regInfo = new XllRegistrationInfo(xllPath);
                        _xllRegistrationInfos[xllPath] = regInfo;
                        //regInfo.Refresh();    // Rather not.... so that we don't even try during the AddIns enumeration... OnInvalidate will lead to Refresh()

                        if (!_isDirty)
                        {
                            _isDirty = true;
                            // This call would case trouble while Excel is shutting down.
                            // CONSIDER: Is there a check we might do.... (and do we in fact get .xll loads during shutdown?)
                            _syncContextExcel.Post(OnInvalidate, null);
                        }

                    }
                }
                else if (notification.Reason == LoaderNotification.Reason.Unloaded)
                {
                    _xllRegistrationInfos.Remove(xllPath);

                    // Not too eager when cleaning up
                    // OnInvalidate();
                }
            }
            Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.ProcessLoadNotification - after lock");
        }

        private string GetXmlPath(string xllPath) => Path.ChangeExtension(xllPath, ".intellisense.xml");

        // Called in macro context
        // Might be implemented by COM AddIns check, or checking loaded Modules with Win32
        // Application.AddIns2 also lists add-ins interactively loaded (Excel 2010+) and has IsOpen property.
        // See: http://blogs.office.com/2010/02/16/migrating-excel-4-macros-to-vba/
        // Alternative on old Excel is DOCUMENTS(2) which lists all loaded .xlm (also .xll?)
        // Alternative, more in line with our update watch, is to enumerate all loaded modules...
        private IEnumerable<string> GetLoadedXllPaths()
        {
            //// DOCUMENT(2) does not seem to include .xll add-ins

            //var loadedDocs = Integration.XlCall.Excel(Integration.XlCall.xlfDocuments, 2) as object[,];
            //if (loadedDocs == null)
            //{
            //    Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.GetLoadedXllPaths - DOCUMENTS(2) failed");
            //    yield break;
            //}
            //for (int i = 0; i < loadedDocs.GetLength(1); i++)
            //{
            //    var docName = loadedDocs[0, i] as string;
            //    if (docName != null && Path.GetExtension(docName) == ".xll")
            //    {
            //        yield return docName;
            //    }
            //}

            //// TODO: Implement properly...
            //dynamic app = ExcelDnaUtil.Application;
            //foreach (var addin in app.AddIns2)
            //{
            //    if (addin.IsOpen && Path.GetExtension(addin.FullName) == ".xll")
            //    {
            //        yield return addin.FullName;
            //    }
            //}

            // Enumerate loaded modules - pick .xll files
            var process = Process.GetCurrentProcess();
            ProcessModuleCollection modules = process.Modules;
            foreach (ProcessModule module in modules)
            {
                var fileName = module.FileName;
                if (Path.GetExtension(fileName) == ".xll")
                {
                    yield return fileName;
                }
            }
        }

        // Must be called on the main thread, in a macro context
        private void OnInvalidate(object _unused_)
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Invalidate?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Dispose");
            _loaderNotification.Dispose();
        }
    }
}
