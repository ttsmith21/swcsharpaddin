using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace NM.Core
{
    /// <summary>
    /// Represents the state of a custom property relative to its original value.
    /// </summary>
    public enum PropertyState
    {
        /// <summary>The property is unchanged.</summary>
        Unchanged = 0,
        /// <summary>The property's value has been modified.</summary>
        Modified = 1,
        /// <summary>The property is new and has been added.</summary>
        Added = 2,
        /// <summary>The property has been marked for deletion.</summary>
        Deleted = 3
    }

    /// <summary>
    /// Defines the data type of a custom property, matching SolidWorks API values.
    /// </summary>
    public enum CustomPropertyType
    {
        /// <summary>A text or string property.</summary>
        Text = 30, // swCustomInfoText
        /// <summary>A numeric (double) property.</summary>
        Number = 3,  // swCustomInfoNumber
        /// <summary>A date/time property.</summary>
        Date = 64    // swCustomInfoDate
    }

    /// <summary>
    /// Manages a collection of custom property data with change tracking and validation.
    /// This class is a pure data container and has no direct dependency on SolidWorks APIs.
    /// </summary>
    public class CustomPropertyData
    {
        #region Private Fields
        private readonly Dictionary<string, object> _properties;
        private readonly Dictionary<string, object> _originalProperties;
        private readonly Dictionary<string, PropertyState> _propertyStates;
        private readonly Dictionary<string, CustomPropertyType> _propertyTypes;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="CustomPropertyData"/> class.
        /// </summary>
        public CustomPropertyData()
        {
            // Initialize dictionaries with case-insensitive key comparison
            var comparer = StringComparer.OrdinalIgnoreCase;
            _properties = new Dictionary<string, object>(comparer);
            _originalProperties = new Dictionary<string, object>(comparer);
            _propertyStates = new Dictionary<string, PropertyState>(comparer);
            _propertyTypes = new Dictionary<string, CustomPropertyType>(comparer);

            InitializeWithDefaults();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets a value indicating whether any properties have been added, modified, or deleted.
        /// </summary>
        public bool IsDirty { get; private set; }

        /// <summary>
        /// Gets or sets the name of the configuration this property set applies to.
        /// An empty string indicates file-specific (global) properties.
        /// </summary>
        public string ConfigurationName { get; set; } = "";

        /// <summary>
        /// Gets or sets a value indicating whether the part is sheet metal.
        /// </summary>
        public bool IsSheetMetal
        {
            get => ParseBool(GetPropertyValue("IsSheetMetal"));
            set => SetPropertyValue("IsSheetMetal", value ? "True" : "False", CustomPropertyType.Text);
        }

        /// <summary>
        /// Gets or sets a value indicating whether the part is a tube.
        /// </summary>
        public bool IsTube
        {
            get => ParseBool(GetPropertyValue("IsTube"));
            set => SetPropertyValue("IsTube", value ? "True" : "False", CustomPropertyType.Text);
        }

        /// <summary>
        /// Gets or sets the thickness of the material in inches.
        /// </summary>
        public double Thickness
        {
            get
            {
                string sValue = GetPropertyValue("Thickness")?.ToString() ?? "";
                if (double.TryParse(sValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                {
                    return val;
                }
                return 0.0;
            }
            set => SetPropertyValue("Thickness", value.ToString("0.####", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        /// <summary>
        /// Gets or sets the thickness of the material in meters.
        /// </summary>
        public double ThicknessInMeters
        {
            get => Thickness * Configuration.Materials.InchesToMeters;
            set => Thickness = value * Configuration.Materials.MetersToInches;
        }

        // Convenience manufacturing properties
        public string OptiMaterial
        {
            get => (GetPropertyValue("OptiMaterial") ?? string.Empty).ToString();
            set => SetPropertyValue("OptiMaterial", value ?? string.Empty, CustomPropertyType.Text);
        }

        public string rbMaterialType
        {
            get => (GetPropertyValue("rbMaterialType") ?? string.Empty).ToString();
            set => SetPropertyValue("rbMaterialType", value ?? string.Empty, CustomPropertyType.Text);
        }

        public string MaterialCategory
        {
            get => (GetPropertyValue("MaterialCategory") ?? string.Empty).ToString();
            set => SetPropertyValue("MaterialCategory", value ?? string.Empty, CustomPropertyType.Text);
        }

        public double MaterialDensity
        {
            get => ParseDouble(GetPropertyValue("MaterialDensity"));
            set => SetPropertyValue("MaterialDensity", value.ToString("0.#####", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double RawWeight
        {
            get => ParseDouble(GetPropertyValue("RawWeight"));
            set => SetPropertyValue("RawWeight", value.ToString("0.###", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double SheetPercent
        {
            get => ParseDouble(GetPropertyValue("SheetPercent"));
            set => SetPropertyValue("SheetPercent", value.ToString("0.####", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double MaterialCostPerLB
        {
            get => ParseDouble(GetPropertyValue("MaterialCostPerLB"));
            set
            {
                var s = value.ToString("0.###", CultureInfo.InvariantCulture);
                SetPropertyValue("MaterialCostPerLB", s, CustomPropertyType.Number);
                // Legacy VBA misspelling compatibility
                SetPropertyValue("MaterailCostPerLB", s, CustomPropertyType.Number);
            }
        }

        public int QuoteQty
        {
            get => (int)Math.Round(ParseDouble(GetPropertyValue("QuoteQty")));
            set => SetPropertyValue("QuoteQty", value.ToString(CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public string Difficulty
        {
            get => (GetPropertyValue("Difficulty") ?? string.Empty).ToString();
            set => SetPropertyValue("Difficulty", value ?? string.Empty, CustomPropertyType.Text);
        }

        public string CuttingType
        {
            get => (GetPropertyValue("CuttingType") ?? string.Empty).ToString();
            set => SetPropertyValue("CuttingType", value ?? string.Empty, CustomPropertyType.Text);
        }

        public string Description
        {
            get => (GetPropertyValue("Description") ?? string.Empty).ToString();
            set => SetPropertyValue("Description", value ?? string.Empty, CustomPropertyType.Text);
        }

        public string Customer
        {
            get => (GetPropertyValue("Customer") ?? string.Empty).ToString();
            set => SetPropertyValue("Customer", value ?? string.Empty, CustomPropertyType.Text);
        }

        // Bend-related properties for cost calculations
        public int BendCount
        {
            get => (int)Math.Round(ParseDouble(GetPropertyValue("BendCount")));
            set => SetPropertyValue("BendCount", value.ToString(CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double LongestBendIn
        {
            get => ParseDouble(GetPropertyValue("LongestBendIn"));
            set => SetPropertyValue("LongestBendIn", value.ToString("0.###", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double MaxBendRadiusIn
        {
            get => ParseDouble(GetPropertyValue("MaxBendRadiusIn"));
            set => SetPropertyValue("MaxBendRadiusIn", value.ToString("0.###", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        public double ArcLengthIn
        {
            get => ParseDouble(GetPropertyValue("ArcLengthIn"));
            set => SetPropertyValue("ArcLengthIn", value.ToString("0.###", CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }

        // Tapped hole count for F220 tapping calculations
        public int TappedHoleCount
        {
            get => (int)Math.Round(ParseDouble(GetPropertyValue("TappedHoleCount")));
            set => SetPropertyValue("TappedHoleCount", value.ToString(CultureInfo.InvariantCulture), CustomPropertyType.Number);
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the value of a property by its name.
        /// </summary>
        /// <param name="propertyName">The case-insensitive name of the property.</param>
        /// <returns>The property value, or null if not found.</returns>
        public object GetPropertyValue(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) return null;
            return _properties.TryGetValue(propertyName, out object value) ? value : null;
        }

        /// <summary>
        /// Sets the value of a property, tracking its state (added, modified).
        /// </summary>
        /// <param name="propertyName">The case-insensitive name of the property.</param>
        /// <param name="value">The new value for the property.</param>
        /// <param name="propertyType">The data type of the property.</param>
        public void SetPropertyValue(string propertyName, object value, CustomPropertyType propertyType = CustomPropertyType.Text)
        {
            const string procName = "CustomPropertyData.SetPropertyValue";
            if (!ValidateString(propertyName, procName, "property name")) return;

            // TODO(vNext): Add more robust validation for numeric and date types.
            if (propertyType == CustomPropertyType.Number && value != null && !double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                ErrorHandler.HandleError(procName, "Non-numeric value for numeric property", null, ErrorHandler.LogLevel.Warning, $"Prop: {propertyName}, Val: {value}");
                return;
            }

            if (_properties.TryGetValue(propertyName, out object oldValue))
            {
                // Property exists, check if value or type has changed
                bool valueChanged = !Equals(oldValue, value);
                bool typeChanged = _propertyTypes.TryGetValue(propertyName, out var oldType) && oldType != propertyType;

                if (valueChanged || typeChanged)
                {
                    _properties[propertyName] = value;
                    _propertyTypes[propertyName] = propertyType;
                    // If it was previously unchanged, mark it as modified. Don't change 'Added' state.
                    if (_propertyStates[propertyName] == PropertyState.Unchanged)
                    {
                        _propertyStates[propertyName] = PropertyState.Modified;
                    }
                    IsDirty = true;
                }
            }
            else
            {
                // New property
                _properties.Add(propertyName, value);
                _originalProperties.Add(propertyName, null); // Original is null for new props
                _propertyStates.Add(propertyName, PropertyState.Added);
                _propertyTypes.Add(propertyName, propertyType);
                IsDirty = true;
            }
        }

        /// <summary>
        /// Marks a property for deletion. The actual removal happens during the write phase.
        /// </summary>
        /// <param name="propertyName">The case-insensitive name of the property to delete.</param>
        /// <returns>True if the property was found and marked for deletion, false otherwise.</returns>
        public bool DeleteProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName) || !_properties.ContainsKey(propertyName))
            {
                return false;
            }

            _propertyStates[propertyName] = PropertyState.Deleted;
            IsDirty = true;
            return true;
        }

        /// <summary>
        /// Clears the values of all existing properties and marks them as modified.
        /// </summary>
        public void ClearAllProperties()
        {
            foreach (var key in _properties.Keys.ToList())
            {
                SetPropertyValue(key, "", _propertyTypes.TryGetValue(key, out var type) ? type : CustomPropertyType.Text);
            }
        }

        /// <summary>
        /// Gets an array of all property names currently in the collection.
        /// </summary>
        /// <returns>An array of property names.</returns>
        public string[] GetAllPropertyNames()
        {
            return _properties.Keys.ToArray();
        }

        /// <summary>
        /// Resets the change tracking, marking all properties as unchanged.
        /// This should be called after successfully writing data to the external source.
        /// </summary>
        public void MarkClean()
        {
            _originalProperties.Clear();
            foreach (var kvp in _properties)
            {
                _originalProperties.Add(kvp.Key, kvp.Value);
                if (_propertyStates.ContainsKey(kvp.Key))
                {
                    _propertyStates[kvp.Key] = PropertyState.Unchanged;
                }
            }
            IsDirty = false;
        }

        /// <summary>
        /// Initializes the property collection with default keys from the application configuration.
        /// </summary>
        public void InitializeWithDefaults()
        {
            _properties.Clear();
            _originalProperties.Clear();
            _propertyStates.Clear();
            _propertyTypes.Clear();

            var initialProps = Configuration.Materials.GetInitialCustomProperties();
            foreach (var propName in initialProps)
            {
                if (!string.IsNullOrWhiteSpace(propName))
                {
                    _properties.Add(propName, "");
                    _originalProperties.Add(propName, "");
                    _propertyStates.Add(propName, PropertyState.Unchanged);
                    _propertyTypes.Add(propName, CustomPropertyType.Text); // Default to text
                }
            }
            IsDirty = false;
        }

        /// <summary>
        /// Gets a dictionary of all properties and their current states.
        /// </summary>
        public IReadOnlyDictionary<string, PropertyState> GetPropertyStates() => _propertyStates;

        /// <summary>
        /// Gets a dictionary of all properties and their data types.
        /// </summary>
        public IReadOnlyDictionary<string, CustomPropertyType> GetPropertyTypes() => _propertyTypes;

        /// <summary>
        /// Gets a dictionary of all current property values.
        /// </summary>
        public IReadOnlyDictionary<string, object> GetProperties() => _properties;

        #endregion

        #region Private Helpers
        private bool ValidateString(string value, string procedureName, string paramName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                ErrorHandler.HandleError(procedureName, $"Invalid {paramName}: cannot be null or empty.", null, ErrorHandler.LogLevel.Warning);
                return false;
            }
            return true;
        }

        private static bool ParseBool(object val)
        {
            if (val == null) return false;
            var s = val.ToString();
            if (bool.TryParse(s, out var b)) return b;
            // accept 1/0
            if (int.TryParse(s, out var i)) return i != 0;
            return string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "y", StringComparison.OrdinalIgnoreCase);
        }

        private static double ParseDouble(object val)
        {
            if (val == null) return 0.0;
            var s = val.ToString();
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
            if (double.TryParse(s, out d)) return d;
            return 0.0;
        }
        #endregion
    }
}
