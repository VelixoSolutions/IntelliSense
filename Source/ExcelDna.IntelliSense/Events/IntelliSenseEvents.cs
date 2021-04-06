#nullable enable

using System;
using System.Collections.Generic;
using System.Reactive;
using System.Reactive.Linq;
using System.Reactive.Subjects;

namespace ExcelDna.IntelliSense
{
    public class IntelliSenseEvents : IIntelliSenseEvents
    {
        public static IntelliSenseEvents Instance = new IntelliSenseEvents();

        private BehaviorSubject<string?> _editedFormula = new BehaviorSubject<string?>(null);

        private BehaviorSubject<string?> _functionName = new BehaviorSubject<string?>(null);

        private BehaviorSubject<(string name, int index)?> _editedArgument = new BehaviorSubject<(string name, int index)?>(null);

        public IObservable<string?> FunctionName => _functionName.DistinctUntilChanged().Publish().RefCount();

        public IObservable<string?> EditedFormula => _editedFormula.DistinctUntilChanged().Publish().RefCount();

        public IObservable<(string name, int index)?> EditedArgument => _editedArgument.DistinctUntilChanged().Publish().RefCount();

        public event EventHandler<CollectingArgumentDescriptionEventArgs>? OnCollectingAdditionalArgumentDescription;

        internal event EventHandler? OnIntellisenseInvalidated;

        public void Invalidate() => OnIntellisenseInvalidated?.Invoke(this, EventArgs.Empty);

        internal void OnEditingFunction(string? functionName) => _functionName.OnNext(functionName);

        internal void OnEditingFormula(string? fullFormula) => _editedFormula.OnNext(fullFormula);

        internal void OnEditingArgument(string? argumentName, int? argumentIndex) 
            => _editedArgument.OnNext(
                    argumentName == null || argumentIndex == null 
                        ? null as (string, int)?
                        : (argumentName, argumentIndex.Value));

        internal IEnumerable<TextLine> RaiseCollectingArgumentDescription()
        {
            var eventArgs = new CollectingArgumentDescriptionEventArgs();

            OnCollectingAdditionalArgumentDescription?.Invoke(null, eventArgs);

            return eventArgs.AdditionalDescriptionLines;
        }
    }
}
