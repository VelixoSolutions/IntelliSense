using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    // This is really the running IntelliSenseServer
    // It brings together:
    // * the UIMonitor which monitors the state of Excel, including the current function prefix,
    // * the IntelliSenseDisplay which presents the pop-ups, and 
    // * the IntelliSenseProviders which figure out what information is available.
    internal class IntelliSenseHelper : IDisposable
    {
        private readonly SynchronizationContext _syncContextMain; // Main thread, not macro context
        private readonly UIMonitor _uiMonitor;  // We want the UIMonitor here, because we might hook up other display enhancements

        // These need to get combined into a UIEnhancement class ....
        private readonly IntelliSenseDisplay _display;
        private readonly List<IIntelliSenseProvider> _providers = new List<IIntelliSenseProvider>();
        // TODO: Others

        public IntelliSenseHelper()
        {
            Logger.Initialization.Verbose("IntelliSenseHelper Constructor Begin");
            _syncContextMain = new WindowsFormsSynchronizationContext();
            _uiMonitor = new UIMonitor(_syncContextMain);
            _display = new IntelliSenseDisplay(_syncContextMain, _uiMonitor);

            _providers = new List<IIntelliSenseProvider>
            {
                new ExcelDnaIntelliSenseProvider(_syncContextMain),
            };

            RegisterIntellisense();
            Logger.Initialization.Verbose("IntelliSenseHelper Constructor End");
        }

        private void RegisterIntellisense()
        {
            foreach (IIntelliSenseProvider provider in _providers)
            {
                provider.Invalidate += Provider_Invalidate;
                provider.Initialize();
                UpdateDisplay(provider);
            }
        }

        // We need to call Refresh on the main thread in a macro context,
        // and then GetFunctionInfos() to update the Display
        private void Provider_Invalidate(object sender, EventArgs e) => RefreshProvider(sender);

        // Must be called on the main thread, in a macro context
        // TODO: Still not sure how to delete / unregister...
        private void RefreshProvider(object providerObj)
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Verbose($"IntelliSenseHelper.RefreshProvider - Begin Refresh for {providerObj.GetType().Name}");
            var provider = (IIntelliSenseProvider)providerObj;
            provider.Refresh();
            UpdateDisplay(provider);
            Logger.Provider.Verbose($"IntelliSenseHelper.RefreshProvider - End");
        }

        private void UpdateDisplay(IIntelliSenseProvider provider)
        {
            IList<FunctionInfo> functionInfos = provider.GetFunctionInfos();
            _display.UpdateFunctionInfos(functionInfos);
        }

        // Must be called on the main thread, in a macro context
        // TODO: Still not sure how to delete / unregister...
        internal void RefreshProviders()
        {
            foreach (IIntelliSenseProvider provider in _providers)
            {
                RefreshProvider(provider);
            }
        }

        // Must run on the main thread
        public void Dispose()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Initialization.Verbose("IntelliSenseHelper Dispose Start");

            foreach (IIntelliSenseProvider provider in _providers)
            {
                provider.Dispose();
            }
            _display.Dispose();
            _uiMonitor.Dispose();

            Logger.Initialization.Verbose("IntelliSenseHelper Dispose End");
        }
    }
}
