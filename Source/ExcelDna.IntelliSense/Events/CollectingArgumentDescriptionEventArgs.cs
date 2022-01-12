#nullable enable

using System;
using System.Collections.Generic;

namespace ExcelDna.IntelliSense
{
    /// <summary>
    /// Indicates a chance to append description lines to the currently processed argument.
    /// The event consumer can choose to append the description with custom text.
    /// </summary>
    public class CollectingArgumentDescriptionEventArgs : EventArgs
    {
        public ICollection<TextLine> AdditionalDescriptionLines { get; } = new List<TextLine>();        
    }
}
