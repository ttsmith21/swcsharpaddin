using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace NM.Core
{
    /// <summary>
    /// Represents a SolidWorks model (part, assembly, or drawing) with state and property management.
    /// Pure data class: contains no SolidWorks COM types or calls.
    /// </summary>
    public class ModelInfo
    {
        #region Enums
        /// <summary>Represents lifecycle state of a model during processing.</summary>
        public enum ModelState
        {
            Idle = 0,
            Opened = 1,
            Validated = 2,
            Problem = 3,
            Processed = 4
        }

        /// <summary>Represents the type of the model document.</summary>
        public enum ModelType
        {
            Unknown = 0,
            Part = 1,
            Assembly = 2,
            Drawing = 3
        }
        #endregion

        #region Nested types
        /// <summary>Represents a single bend parameter set.</summary>
        public class BendParameter
        {
            /// <summary>Bend radius (meters).</summary>
            public double Radius { get; set; }
            /// <summary>Thickness (meters).</summary>
            public double Thickness { get; set; }
            /// <summary>K-factor (0..1).</summary>
            public double KFactor { get; set; }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates an empty ModelInfo; call Initialize to set file path and configuration.
        /// </summary>
        public ModelInfo()
        {
            CustomProperties = new CustomPropertyData();
            _bendParameters = new List<BendParameter>();
            CurrentState = ModelState.Idle;
            Type = ModelType.Unknown;
            KFactor = 0.5; // default
            ConfigurationName = string.Empty;
        }
        #endregion

        #region Data properties
        /// <summary>Full path to the model file.</summary>
        public string FilePath { get; private set; } = string.Empty;
        /// <summary>Model type derived from file extension.</summary>
        public ModelType Type { get; private set; } = ModelType.Unknown;
        /// <summary>Current processing state.</summary>
        public ModelState CurrentState { get; set; } = ModelState.Idle;
        /// <summary>Active configuration name; empty for global.</summary>
        public string ConfigurationName
        {
            get => _configurationName;
            set
            {
                _configurationName = value ?? string.Empty;
                // Note: CustomPropertyData currently stores config name in this class; no direct coupling required yet.
                IsDirty = true;
            }
        }
        private string _configurationName = string.Empty;
        /// <summary>Problem description when state is Problem.</summary>
        public string ProblemDescription { get; set; } = string.Empty;
        /// <summary>Indicates there are pending changes.</summary>
        public bool IsDirty { get; set; }
        /// <summary>Indicates the model needs fixing (set by analyzers).</summary>
        public bool NeedsFix { get; set; }

        // Manufacturing flags
        /// <summary>True if part is sheet metal. Proxies to custom properties.</summary>
        public bool IsSheetMetal
        {
            get => CustomProperties.IsSheetMetal;
            set { CustomProperties.IsSheetMetal = value; IsDirty = true; }
        }
        /// <summary>True if part is flattened (for downstream checks).</summary>
        public bool IsFlattened { get; set; }
        /// <summary>Indicates whether a previous insert operation succeeded.</summary>
        public bool InsertSuccessful { get; set; }

        // Thickness
        /// <summary>Thickness in inches. Proxies to custom properties.</summary>
        public double ThicknessInInches
        {
            get => CustomProperties.Thickness;
            set { CustomProperties.Thickness = value; IsDirty = true; }
        }
        /// <summary>Thickness in meters. Proxies to custom properties.</summary>
        public double ThicknessInMeters
        {
            get => CustomProperties.ThicknessInMeters;
            set { CustomProperties.ThicknessInMeters = value; IsDirty = true; }
        }

        /// <summary>Bend radius (meters).</summary>
        public double BendRadius
        {
            get => _bendRadius;
            set
            {
                if (value <= 0)
                {
                    ErrorHandler.HandleError("ModelInfo.BendRadius", $"Invalid bend radius: {value}", null, ErrorHandler.LogLevel.Warning);
                    return;
                }
                _bendRadius = value; IsDirty = true;
            }
        }
        private double _bendRadius;

        /// <summary>K-factor (0..1).</summary>
        public double KFactor
        {
            get => _kFactor;
            set
            {
                if (value < 0 || value > 1)
                {
                    ErrorHandler.HandleError("ModelInfo.KFactor", $"Invalid K-factor: {value}", null, ErrorHandler.LogLevel.Warning);
                    return;
                }
                _kFactor = value; IsDirty = true;
            }
        }
        private double _kFactor;

        // Mass/volume tracking
        /// <summary>Initial mass (kg).</summary>
        public double InitialMass
        {
            get => _initialMass;
            set { if (value >= 0) { _initialMass = value; IsDirty = true; } else ErrorHandler.HandleError("ModelInfo.InitialMass", $"Invalid InitialMass: {value}", null, ErrorHandler.LogLevel.Warning); }
        }
        private double _initialMass;
        /// <summary>Initial volume (m^3).</summary>
        public double InitialVolume
        {
            get => _initialVolume;
            set { if (value >= 0) { _initialVolume = value; IsDirty = true; } else ErrorHandler.HandleError("ModelInfo.InitialVolume", $"Invalid InitialVolume: {value}", null, ErrorHandler.LogLevel.Warning); }
        }
        private double _initialVolume;
        /// <summary>Final volume (m^3).</summary>
        public double FinalVolume
        {
            get => _finalVolume;
            set { if (value >= 0) { _finalVolume = value; IsDirty = true; } else ErrorHandler.HandleError("ModelInfo.FinalVolume", $"Invalid FinalVolume: {value}", null, ErrorHandler.LogLevel.Warning); }
        }
        private double _finalVolume;

        /// <summary>Custom properties container for this model.</summary>
        public CustomPropertyData CustomProperties { get; }
        #endregion

        #region Bend parameters
        private readonly List<BendParameter> _bendParameters;
        /// <summary>Adds a bend parameter set (meters, meters, 0..1).</summary>
        public void AddBendParameter(double radius, double thickness, double kFactor)
        {
            if (radius <= 0 || thickness <= 0 || kFactor < 0 || kFactor > 1)
            {
                ErrorHandler.HandleError("ModelInfo.AddBendParameter", $"Invalid input - Thickness={thickness}, Radius={radius}, KFactor={kFactor}", null, ErrorHandler.LogLevel.Warning);
                return;
            }
            _bendParameters.Add(new BendParameter { Radius = radius, Thickness = thickness, KFactor = kFactor });
            IsDirty = true;
        }
        /// <summary>Returns a copy of the bend parameter list.</summary>
        public List<BendParameter> GetBendParameters() => new List<BendParameter>(_bendParameters);
        /// <summary>Clears all bend parameters.</summary>
        public void ClearBendParameters() { _bendParameters.Clear(); IsDirty = true; }
        #endregion

        #region Initialization and validation
        /// <summary>
        /// Initialize the model with a file path and optional configuration name.
        /// </summary>
        /// <param name="filePath">Full path to model file.</param>
        /// <param name="configurationName">Configuration name (optional).</param>
        public void Initialize(string filePath, string configurationName = "")
        {
            const string procName = "ModelInfo.Initialize";
            ErrorHandler.PushCallStack(procName);
            try
            {
                // Allow unsaved/virtual documents: path can be empty or non-existent.
                // We still store the provided value and expose IsValidFile for consumers.
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    ErrorHandler.HandleError(procName, "Model initialized with empty path (unsaved document)", null, ErrorHandler.LogLevel.Warning);
                }
                else if (!File.Exists(filePath))
                {
                    // Don't throw; just warn and continue so in-memory docs can work.
                    ErrorHandler.HandleError(procName, $"File does not exist: {filePath}", null, ErrorHandler.LogLevel.Warning);
                }

                if (!ValidateDependencies())
                {
                    ErrorHandler.HandleError(procName, "Missing required dependencies", null, ErrorHandler.LogLevel.Error, $"Path: {filePath}");
                    throw new InvalidOperationException("Missing dependencies");
                }

                FilePath = filePath ?? string.Empty;
                ConfigurationName = configurationName ?? string.Empty;
                Type = DetermineModelType(filePath);
                CurrentState = ModelState.Idle;
                IsDirty = false;
                ProblemDescription = string.Empty;
                CustomProperties.InitializeWithDefaults();
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private ModelType DetermineModelType(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt": return ModelType.Part;
                case ".sldasm": return ModelType.Assembly;
                case ".slddrw": return ModelType.Drawing;
                default: return ModelType.Unknown;
            }
        }

        private bool ValidateDependencies()
        {
            // Basic presence check for config constants. In .NET this will always be available if built.
            try
            {
                var _ = Configuration.Materials.MetersToInches;
                _ = Configuration.Materials.InchesToMeters;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ModelInfo.ValidateDependencies", ex.Message, ex);
                return false;
            }
        }
        #endregion

        #region State helpers
        /// <summary>Sets model state.</summary>
        public void SetState(ModelState newState) => CurrentState = newState;
        /// <summary>Marks problem state and records description.</summary>
        public void MarkProblem(string description)
        {
            CurrentState = ModelState.Problem;
            ProblemDescription = description ?? string.Empty;
        }
        #endregion

        #region File info helpers (no IO)
        /// <summary>True if FilePath is non-empty and exists on disk.</summary>
        public bool IsValidFile => !string.IsNullOrWhiteSpace(FilePath) && File.Exists(FilePath);
        /// <summary>File name without directory.</summary>
        public string FileName => string.IsNullOrWhiteSpace(FilePath) ? string.Empty : Path.GetFileName(FilePath);
        /// <summary>Extension including the leading dot.</summary>
        public string FileExtension => string.IsNullOrWhiteSpace(FilePath) ? string.Empty : Path.GetExtension(FilePath);
        #endregion
    }
}
