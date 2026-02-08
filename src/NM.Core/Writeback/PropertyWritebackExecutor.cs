using System;
using System.Collections.Generic;
using NM.Core.Reconciliation.Models;
using NM.Core.Writeback.Models;

namespace NM.Core.Writeback
{
    /// <summary>
    /// Applies approved property suggestions to the CustomPropertyData cache.
    /// After calling ApplyToCache, the caller uses the existing
    /// CustomPropertiesService.WritePending() to flush changes to SolidWorks.
    ///
    /// Flow: PropertySuggestions → ApplyToCache(cache) → WritePending(model)
    /// </summary>
    public sealed class PropertyWritebackExecutor
    {
        /// <summary>
        /// Applies approved property suggestions to the custom property cache.
        /// Only writes suggestions that pass validation. Skips null/empty values
        /// and properties that already have the suggested value.
        /// </summary>
        /// <param name="approvedSuggestions">Suggestions the user has approved in the UI wizard.</param>
        /// <param name="cache">The model's custom property cache (with change tracking).</param>
        /// <returns>Result tracking what was applied, skipped, or failed.</returns>
        public WritebackResult ApplyToCache(
            IList<PropertySuggestion> approvedSuggestions,
            CustomPropertyData cache)
        {
            if (approvedSuggestions == null)
                throw new ArgumentNullException(nameof(approvedSuggestions));
            if (cache == null)
                throw new ArgumentNullException(nameof(cache));

            var result = new WritebackResult();

            foreach (var suggestion in approvedSuggestions)
            {
                if (string.IsNullOrWhiteSpace(suggestion.PropertyName))
                {
                    result.Skipped.Add(new WritebackEntry
                    {
                        PropertyName = suggestion.PropertyName ?? "(null)",
                        NewValue = suggestion.Value,
                        Status = WritebackStatus.Skipped,
                        Reason = "Property name is empty or null"
                    });
                    continue;
                }

                // Read current value from cache
                string currentValue = cache.GetPropertyValue(suggestion.PropertyName)?.ToString() ?? string.Empty;
                string newValue = suggestion.Value ?? string.Empty;

                // Skip if value is already the same
                if (string.Equals(currentValue, newValue, StringComparison.OrdinalIgnoreCase))
                {
                    result.Skipped.Add(new WritebackEntry
                    {
                        PropertyName = suggestion.PropertyName,
                        OldValue = currentValue,
                        NewValue = newValue,
                        Status = WritebackStatus.Skipped,
                        Reason = "Value already matches"
                    });
                    continue;
                }

                try
                {
                    // Determine property type (most are text; numeric if it looks numeric)
                    var propType = InferPropertyType(suggestion.PropertyName, newValue);

                    cache.SetPropertyValue(suggestion.PropertyName, newValue, propType);

                    result.Applied.Add(new WritebackEntry
                    {
                        PropertyName = suggestion.PropertyName,
                        OldValue = currentValue,
                        NewValue = newValue,
                        Status = WritebackStatus.Applied,
                        Category = suggestion.Category
                    });
                }
                catch (Exception ex)
                {
                    result.Failed.Add(new WritebackEntry
                    {
                        PropertyName = suggestion.PropertyName,
                        OldValue = currentValue,
                        NewValue = newValue,
                        Status = WritebackStatus.Failed,
                        Reason = ex.Message
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Applies a single property value to the cache. Useful for conflict
        /// resolutions where the user picks a value interactively.
        /// </summary>
        public WritebackEntry ApplySingle(
            string propertyName, string value, CustomPropertyData cache)
        {
            if (cache == null)
                throw new ArgumentNullException(nameof(cache));

            string currentValue = cache.GetPropertyValue(propertyName)?.ToString() ?? string.Empty;

            try
            {
                var propType = InferPropertyType(propertyName, value ?? string.Empty);
                cache.SetPropertyValue(propertyName, value ?? string.Empty, propType);

                return new WritebackEntry
                {
                    PropertyName = propertyName,
                    OldValue = currentValue,
                    NewValue = value ?? string.Empty,
                    Status = WritebackStatus.Applied
                };
            }
            catch (Exception ex)
            {
                return new WritebackEntry
                {
                    PropertyName = propertyName,
                    OldValue = currentValue,
                    NewValue = value ?? string.Empty,
                    Status = WritebackStatus.Failed,
                    Reason = ex.Message
                };
            }
        }

        /// <summary>
        /// Infer property type from name and value.
        /// Setup/run times and operation numbers are numeric; everything else is text.
        /// </summary>
        private static CustomPropertyType InferPropertyType(string propertyName, string value)
        {
            if (string.IsNullOrEmpty(propertyName))
                return CustomPropertyType.Text;

            // Numeric properties: setup times (_S), run times (_R), operation numbers
            if (propertyName.EndsWith("_S", StringComparison.OrdinalIgnoreCase) ||
                propertyName.EndsWith("_R", StringComparison.OrdinalIgnoreCase) ||
                propertyName.Equals("OtherOP", StringComparison.OrdinalIgnoreCase) ||
                propertyName.StartsWith("Other_OP", StringComparison.OrdinalIgnoreCase))
            {
                if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out _))
                {
                    return CustomPropertyType.Number;
                }
            }

            return CustomPropertyType.Text;
        }
    }
}
