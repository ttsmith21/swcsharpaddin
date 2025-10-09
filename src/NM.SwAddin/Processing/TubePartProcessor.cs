using System;
using NM.Core;
using NM.Core.Processing;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Processing
{
    /// <summary>
    /// IPartProcessor adapter that delegates to NM.Core.SimpleTubeProcessor for detection and processing.
    /// Writes minimal tube properties for downstream consumers.
    /// </summary>
    public sealed class TubePartProcessor : IPartProcessor
    {
        public ProcessorType Type => ProcessorType.Tube;

        private readonly SimpleTubeProcessor _core = new SimpleTubeProcessor();

        public bool CanProcess(IModelDoc2 model)
        {
            if (model == null) return false;
            return _core.CanProcess(model);
        }

        public ProcessingResult Process(IModelDoc2 model, ModelInfo info, ProcessingOptions options)
        {
            try
            {
                if (model == null || info == null)
                    return ProcessingResult.Fail("Invalid inputs", NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.Fatal);

                // Extract geometry first
                var geom = _core.TryGetGeometry(model);
                if (geom == null)
                {
                    return ProcessingResult.Fail("Tube geometry not detected", NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.GeometryValidation);
                }

                // Mark properties
                info.CustomProperties.IsTube = true;
                info.CustomProperties.SetPropertyValue("TubeOD", geom.OuterDiameter.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                info.CustomProperties.SetPropertyValue("TubeWall", geom.WallThickness.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                info.CustomProperties.SetPropertyValue("TubeLength", geom.Length.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);

                // Call core processor (no-op for now, logs only)
                var ok = _core.Process(model, options ?? new ProcessingOptions());
                if (!ok)
                    return ProcessingResult.Fail("Core tube processing failed", NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.ProcessingError);

                return ProcessingResult.Ok(Type.ToString());
            }
            catch (Exception ex)
            {
                return ProcessingResult.Fail("Tube processing exception: " + ex.Message, NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.Fatal);
            }
        }
    }
}
