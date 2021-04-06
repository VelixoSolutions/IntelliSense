#nullable enable

using System;

namespace ExcelDna.IntelliSense
{
    public interface IIntelliSenseEvents
    {
        public event EventHandler<CollectingArgumentDescriptionEventArgs>? OnCollectingAdditionalArgumentDescription;

        public IObservable<string?> FunctionName { get; }

        public IObservable<string?> EditedFormula { get; }

        public IObservable<(string name, int index)?> EditedArgument { get; }

        /// <summary>
        /// Invoked by the client when new, perhaps asynchronous, IntelliSense data becomes
        /// available in the current context. This will signal ExcelDNA to refresh the IntelliSense tooltip.
        /// Eventually, <see cref="OnCollectingAdditionalArgumentDescription"/> will be invoked again, where
        /// the client will have a chance to actually pass in the new description lines.
        /// </summary>
        public void Invalidate();
    }
}
