using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    // written by LLM, 2026-11-14
    public static class WikiTraccsPageCreationFields
    {
        public const string ItemIdFieldName = "WT_Out_ItemId";
        public const string CreatedUtcMsFieldName = "WT_Internal_FirstCreatedUtcMs";

        /// Returns only the creation-time field as a double for a SharePoint Number column.
        /// Null means that no integer creation time was supplied.
        /// This subset bypasses PnP's field-type conversion; all other fields use that conversion.
        internal static IDictionary<string, object> GetTypedOverridesThatBypassPnPFieldConversion(IDictionary<string, string> fieldValues)
        {
            if (fieldValues.TryGetValue(CreatedUtcMsFieldName, out var text)
                && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var milliseconds))
            {
                return new Dictionary<string, object> { [CreatedUtcMsFieldName] = (double)milliseconds };
            }
            return null;
        }

        /// Adds creation fields to the metadata that page provisioning already writes.
        // covered by TestPrecreatedPages
        public static void FillTemplate(IDictionary<string, string> fieldValues, int itemId, long createdUtcMs,
            IReadOnlyCollection<string> availableFields)
        {
            if (createdUtcMs <= 0 || itemId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(createdUtcMs));
            }
            // written by LLM, 2026-11-15
            var creationFields = new[]
            {
                (Name: ItemIdFieldName, Value: itemId.ToString(CultureInfo.InvariantCulture)),
                (Name: CreatedUtcMsFieldName, Value: createdUtcMs.ToString(CultureInfo.InvariantCulture))
            };
            foreach (var (name, value) in creationFields)
            {
                if (!availableFields.Contains(name, StringComparer.Ordinal)
                    || (fieldValues.TryGetValue(name, out var existing) && !string.IsNullOrWhiteSpace(existing)))
                {
                    continue;
                }
                fieldValues[name] = value;
            }
        }
    }
}
