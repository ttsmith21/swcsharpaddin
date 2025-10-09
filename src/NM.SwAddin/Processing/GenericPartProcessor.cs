using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    public sealed class GenericPartProcessor : IPartProcessor
    {
        public ProcessorType Type => ProcessorType.Generic;

        public bool CanProcess(IModelDoc2 model)
        {
            if (model == null) return false;
            return model.GetType() == (int)swDocumentTypes_e.swDocPART;
        }

        public ProcessingResult Process(IModelDoc2 model, NM.Core.ModelInfo info, ProcessingOptions options)
        {
            try
            {
                if (model == null || info == null)
                    return ProcessingResult.Fail("Invalid inputs", NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.Fatal);

                // Basic mass properties
                var mp = model.Extension?.CreateMassProperty() as IMassProperty;
                if (mp != null)
                {
                    info.CustomProperties.SetPropertyValue("RawWeight", (mp.Mass * NM.Core.Configuration.Materials.KgToLbs).ToString("0.###"), NM.Core.CustomPropertyType.Number);
                }

                // Simple feature count
                int holes = 0, fillets = 0, total = 0;
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    total++;
                    var tn = feat.GetTypeName2() ?? string.Empty;
                    if (tn.IndexOf("Hole", StringComparison.OrdinalIgnoreCase) >= 0) holes++;
                    if (tn.IndexOf("Fillet", StringComparison.OrdinalIgnoreCase) >= 0) fillets++;
                    feat = feat.GetNextFeature() as IFeature;
                }
                info.CustomProperties.SetPropertyValue("FeatureCount", total.ToString(), NM.Core.CustomPropertyType.Number);
                info.CustomProperties.SetPropertyValue("HoleCount", holes.ToString(), NM.Core.CustomPropertyType.Number);
                info.CustomProperties.SetPropertyValue("FilletCount", fillets.ToString(), NM.Core.CustomPropertyType.Number);

                return ProcessingResult.Ok(Type.ToString());
            }
            catch (Exception ex)
            {
                return ProcessingResult.Fail("Generic processing failed: " + ex.Message, NM.Core.ProblemParts.ProblemPartManager.ProblemCategory.GeometryValidation);
            }
        }
    }
}
