using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using NM.Core;
using NM.Core.Reconciliation.Models;
using NM.Core.Writeback;
using NM.Core.Writeback.Models;

namespace NM.Core.Tests.Writeback
{
    public class PropertyWritebackExecutorTests
    {
        private readonly PropertyWritebackExecutor _executor = new PropertyWritebackExecutor();

        private static CustomPropertyData MakeCache(params (string name, string value)[] props)
        {
            var cache = new CustomPropertyData();
            foreach (var (name, value) in props)
            {
                cache.SetPropertyValue(name, value, CustomPropertyType.Text);
            }
            cache.MarkClean(); // Reset to baseline
            return cache;
        }

        // --- Basic apply ---

        [Fact]
        public void ApplyToCache_WritesDescriptionToCache()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion
                {
                    PropertyName = "Description",
                    Value = "MOUNTING BRACKET",
                    Category = PropertyCategory.Identity
                }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.True(result.Success);
            Assert.Single(result.Applied);
            Assert.Equal("MOUNTING BRACKET", cache.GetPropertyValue("Description")?.ToString());
        }

        [Fact]
        public void ApplyToCache_WritesMultipleProperties()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "Print", Value = "12345-01" },
                new PropertySuggestion { PropertyName = "Description", Value = "BRACKET" },
                new PropertySuggestion { PropertyName = "Revision", Value = "C" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Equal(3, result.Applied.Count);
            Assert.Equal("12345-01", cache.GetPropertyValue("Print")?.ToString());
            Assert.Equal("BRACKET", cache.GetPropertyValue("Description")?.ToString());
            Assert.Equal("C", cache.GetPropertyValue("Revision")?.ToString());
        }

        [Fact]
        public void ApplyToCache_OverwritesExistingValue()
        {
            var cache = MakeCache(("Description", "OLD DESC"));
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "Description", Value = "NEW DESC" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Single(result.Applied);
            Assert.Equal("OLD DESC", result.Applied[0].OldValue);
            Assert.Equal("NEW DESC", result.Applied[0].NewValue);
            Assert.True(result.Applied[0].IsChanged);
        }

        // --- Skipping ---

        [Fact]
        public void ApplyToCache_SkipsWhenValueAlreadyMatches()
        {
            var cache = MakeCache(("Description", "SAME"));
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "Description", Value = "SAME" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Empty(result.Applied);
            Assert.Single(result.Skipped);
            Assert.Equal("Value already matches", result.Skipped[0].Reason);
        }

        [Fact]
        public void ApplyToCache_SkipsEmptyPropertyName()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "", Value = "test" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Empty(result.Applied);
            Assert.Single(result.Skipped);
        }

        [Fact]
        public void ApplyToCache_SkipsNullPropertyName()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = null, Value = "test" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Empty(result.Applied);
            Assert.Single(result.Skipped);
        }

        // --- Routing properties ---

        [Fact]
        public void ApplyToCache_WritesF210CheckboxAndNote()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "F210", Value = "1" },
                new PropertySuggestion { PropertyName = "F210_RN", Value = "BREAK ALL EDGES" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Equal(2, result.Applied.Count);
            Assert.Equal("1", cache.GetPropertyValue("F210")?.ToString());
            Assert.Equal("BREAK ALL EDGES", cache.GetPropertyValue("F210_RN")?.ToString());
        }

        [Fact]
        public void ApplyToCache_WritesOtherWCSlotProperties()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "OtherWC_CB", Value = "1" },
                new PropertySuggestion { PropertyName = "OtherOP", Value = "60" },
                new PropertySuggestion { PropertyName = "Other_WC", Value = "F400" },
                new PropertySuggestion { PropertyName = "Other_RN", Value = "WELD PER DWG" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Equal(4, result.Applied.Count);
            Assert.Equal("F400", cache.GetPropertyValue("Other_WC")?.ToString());
        }

        [Fact]
        public void ApplyToCache_WritesOP20WorkCenter()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "OP20", Value = "F115" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Single(result.Applied);
            Assert.Equal("F115", cache.GetPropertyValue("OP20")?.ToString());
        }

        // --- Numeric type inference ---

        [Fact]
        public void ApplyToCache_InfersNumericTypeForSetupTime()
        {
            var cache = MakeCache();
            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "Other_S", Value = "15" }
            };

            var result = _executor.ApplyToCache(suggestions, cache);

            Assert.Single(result.Applied);
            Assert.Equal("15", cache.GetPropertyValue("Other_S")?.ToString());
        }

        // --- Edge cases ---

        [Fact]
        public void ApplyToCache_EmptyList_ReturnsEmptyResult()
        {
            var cache = MakeCache();
            var result = _executor.ApplyToCache(new List<PropertySuggestion>(), cache);

            Assert.True(result.Success);
            Assert.Equal(0, result.TotalProcessed);
        }

        [Fact]
        public void ApplyToCache_ThrowsOnNullSuggestions()
        {
            var cache = MakeCache();
            Assert.Throws<ArgumentNullException>(() =>
                _executor.ApplyToCache(null, cache));
        }

        [Fact]
        public void ApplyToCache_ThrowsOnNullCache()
        {
            var suggestions = new List<PropertySuggestion>();
            Assert.Throws<ArgumentNullException>(() =>
                _executor.ApplyToCache(suggestions, null));
        }

        [Fact]
        public void ApplyToCache_MarkesCacheDirty()
        {
            var cache = MakeCache();
            Assert.False(cache.IsDirty);

            var suggestions = new List<PropertySuggestion>
            {
                new PropertySuggestion { PropertyName = "Description", Value = "TEST" }
            };

            _executor.ApplyToCache(suggestions, cache);
            Assert.True(cache.IsDirty);
        }

        // --- ApplySingle ---

        [Fact]
        public void ApplySingle_WritesValue()
        {
            var cache = MakeCache();
            var entry = _executor.ApplySingle("Description", "BRACKET", cache);

            Assert.Equal(WritebackStatus.Applied, entry.Status);
            Assert.Equal("BRACKET", cache.GetPropertyValue("Description")?.ToString());
        }

        [Fact]
        public void ApplySingle_TracksOldValue()
        {
            var cache = MakeCache(("Description", "OLD"));
            var entry = _executor.ApplySingle("Description", "NEW", cache);

            Assert.Equal("OLD", entry.OldValue);
            Assert.Equal("NEW", entry.NewValue);
        }

        // --- Result summary ---

        [Fact]
        public void WritebackResult_Summary_FormatsCorrectly()
        {
            var result = new WritebackResult();
            result.Applied.Add(new WritebackEntry { PropertyName = "A", OldValue = "", NewValue = "1", Status = WritebackStatus.Applied });
            result.Applied.Add(new WritebackEntry { PropertyName = "B", OldValue = "", NewValue = "2", Status = WritebackStatus.Applied });
            result.Skipped.Add(new WritebackEntry { PropertyName = "C", Status = WritebackStatus.Skipped });

            Assert.Equal("2 applied, 1 skipped, 0 failed", result.Summary);
            Assert.True(result.Success);
            Assert.Equal(3, result.TotalProcessed);
        }
    }
}
